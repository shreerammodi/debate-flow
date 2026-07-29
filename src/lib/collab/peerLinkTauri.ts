/**
 * The desktop adapter: the PeerLink port over the shell's collab commands.
 *
 * The shell carries bytes and nothing else. Everything about what the peers
 * say lives above this line and is proven against the in-memory transport, so
 * this module's whole job is to keep connection ids straight and to turn each
 * line of JSON back into a wire message.
 */

import { listenHere } from "@/lib/windowEvents";

import {
    parseWireMessage,
    type PeerConn,
    type PeerLink,
    type PeerLinkConfig,
    type WireMessage,
} from "./peerLink";

/** The slice of Tauri this needs, injected so the suite can drive it. */
export interface TauriBridge {
    invoke(cmd: string, args: Record<string, unknown>): Promise<unknown>;
    listen(event: string, cb: (payload: unknown) => void): Promise<() => void>;
}

type ConnectionKind = "direct" | "relayed";

interface PeerPayload {
    connId: string;
    endpointId: string;
    connectionType: ConnectionKind;
}

function asPeer(payload: unknown): PeerPayload | null {
    if (payload === null || typeof payload !== "object") return null;
    const p = payload as Partial<PeerPayload>;
    if (typeof p.connId !== "string" || typeof p.endpointId !== "string") return null;
    return {
        connId: p.connId,
        endpointId: p.endpointId,
        // Anything the shell did not call direct is disclosed as relayed.
        connectionType: p.connectionType === "direct" ? "direct" : "relayed",
    };
}

function asMessage(payload: unknown): { connId: string; msg: WireMessage } | null {
    if (payload === null || typeof payload !== "object") return null;
    const p = payload as { connId?: unknown; payload?: unknown };
    if (typeof p.connId !== "string" || typeof p.payload !== "string") return null;
    let raw: unknown;
    try {
        raw = JSON.parse(p.payload);
    } catch {
        return null;
    }
    // A line that does not conform to its variant is a peer speaking a
    // language this build does not know, or one probing for a field the
    // protocol dereferences without asking. Dropping it beats tearing the link
    // down, and beats letting it reach the handshake.
    const msg = parseWireMessage(raw);
    return msg ? { connId: p.connId, msg } : null;
}

async function defaultBridge(): Promise<TauriBridge> {
    // Dynamic so the browser bundle never pulls in Tauri's JS API, matching
    // how every other desktop touchpoint is gated.
    const core = await import("@tauri-apps/api/core");
    return {
        invoke: (cmd, args) => core.invoke(cmd, args),
        // A session belongs to the window that started it. `listenHere` is what
        // keeps a second window from seeing this one's traffic, or adopting an
        // inbound peer that dialled it.
        listen: (name, cb) => listenHere(name, cb),
    };
}

interface Held {
    conn: PeerConn;
    kind: ConnectionKind;
    onMessage: ((m: WireMessage) => void)[];
    onClose: (() => void)[];
    open: boolean;
    /**
     * The claim this window has in flight, which every write queues behind.
     *
     * The shell refuses a write to a connection another window owns, and the
     * ack admitting a peer goes out the moment the claim does. Two invokes are
     * two IPC requests and nothing orders them, so the ack could reach the
     * shell before the claim and be refused as somebody else's. Null once
     * there is nothing to wait for, so an ordinary send still reaches the
     * shell in the same turn.
     */
    claiming: Promise<unknown> | null;
}

export async function createPeerLink(
    config: PeerLinkConfig,
    bridge?: TauriBridge,
): Promise<PeerLink> {
    const shell = bridge ?? (await defaultBridge());

    const endpointId = (await shell.invoke("collab_start", {
        relay: config.relay,
        mdns: config.discovery === "mdns",
    })) as string;

    const held = new Map<string, Held>();
    let onPeer: ((conn: PeerConn) => void) | null = null;
    const unlisten: (() => void)[] = [];
    /** This link's one hold on the shell's endpoint, spent exactly once. */
    let stopped = false;

    /**
     * Forgets a connection the shell no longer holds, and tells whoever was
     * using it. The shell is the authority: once it has dropped a connection
     * nothing can be sent over it again, so this is a close in every sense but
     * the one command it does not need to issue.
     */
    function dropConn(connId: string): void {
        const entry = held.get(connId);
        if (!entry?.open) return;
        entry.open = false;
        held.delete(connId);
        for (const cb of entry.onClose) cb();
    }

    function makeConn(connId: string, remote: string, kind: ConnectionKind): PeerConn {
        const entry: Held = {
            open: true,
            kind,
            onMessage: [],
            onClose: [],
            claiming: null,
            conn: {
                id: remote,
                connectionType: () => entry.kind,
                claim() {
                    if (!entry.open || entry.claiming) return;
                    // A claim the shell refuses means another window admitted
                    // this peer first, so the connection is theirs and this
                    // side has nothing left to say on it. Dropping it is the
                    // same answer a refused send gets, for the same reason.
                    entry.claiming = shell.invoke("collab_claim", { connId });
                    void entry.claiming.catch(() => dropConn(connId));
                },
                send(msg) {
                    if (!entry.open) return;
                    const write = () =>
                        shell.invoke("collab_send", { connId, payload: JSON.stringify(msg) });
                    // The shell refuses a send for one reason: it is not
                    // holding this connection. A peer that quit and an endpoint
                    // that stopped both land here, neither is retryable, and a
                    // peer going away is ordinary. So the link is dropped
                    // rather than left claiming to be up, and nothing about it
                    // reaches the debater as an error.
                    void (entry.claiming ? entry.claiming.then(write) : write()).catch(() =>
                        dropConn(connId),
                    );
                },
                onMessage(cb) {
                    entry.onMessage.push(cb);
                },
                onClose(cb) {
                    entry.onClose.push(cb);
                },
                close() {
                    if (!entry.open) return;
                    entry.open = false;
                    const hangUp = () => shell.invoke("collab_close", { connId });
                    // Behind the claim for the reason a write is: the shell
                    // refuses a hang-up from a window that does not own the
                    // connection, and a close that overtook this window's own
                    // claim would be refused as somebody else's. Already gone
                    // on the shell's side is the ordinary way a close races a
                    // peer that hung up first, so nothing is made of either.
                    void (entry.claiming ? entry.claiming.then(hangUp) : hangUp()).catch(() => {});
                    held.delete(connId);
                    for (const cb of entry.onClose) cb();
                },
            },
        };
        held.set(connId, entry);
        return entry.conn;
    }

    unlisten.push(
        await shell.listen("collab:peer", (payload) => {
            const peer = asPeer(payload);
            if (!peer) return;
            // Only an inbound connection is announced. A dial reports its own
            // path in its return value, so nothing here can race it.
            if (held.has(peer.connId)) return;
            onPeer?.(makeConn(peer.connId, peer.endpointId, peer.connectionType));
        }),
    );

    unlisten.push(
        await shell.listen("collab:message", (payload) => {
            const parsed = asMessage(payload);
            if (!parsed) return;
            const entry = held.get(parsed.connId);
            if (!entry) return;
            for (const cb of entry.onMessage) cb(parsed.msg);
        }),
    );

    unlisten.push(
        await shell.listen("collab:closed", (payload) => {
            if (payload === null || typeof payload !== "object") return;
            const p = payload as { connId?: unknown };
            if (typeof p.connId !== "string") return;
            dropConn(p.connId);
        }),
    );

    return {
        async endpointId() {
            return endpointId;
        },

        async listen(cb) {
            onPeer = cb;
        },

        async dial(target) {
            const result = (await shell.invoke("collab_dial", { endpointId: target })) as {
                connId: string;
                connectionType: ConnectionKind;
            };
            return makeConn(result.connId, target, result.connectionType);
        },

        async stop() {
            // Once, whatever the caller does. The shell refcounts the endpoint
            // so two links can share one bind, and a second stop from this
            // link would spend a hold it does not have and pull the endpoint
            // out from under the other one.
            if (stopped) return;
            stopped = true;
            for (const un of unlisten) un();
            unlisten.length = 0;
            onPeer = null;
            for (const entry of held.values()) entry.open = false;
            held.clear();
            await shell.invoke("collab_stop", {});
        },
    };
}
