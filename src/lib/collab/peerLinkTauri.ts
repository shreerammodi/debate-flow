/**
 * The desktop adapter: the PeerLink port over the shell's collab commands.
 *
 * The shell carries bytes and nothing else. Everything about what the peers
 * say lives above this line and is proven against the in-memory transport, so
 * this module's whole job is to keep connection ids straight and to turn each
 * line of JSON back into a wire message.
 */

import type { PeerConn, PeerLink, PeerLinkConfig, WireMessage } from "./peerLink";

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
    try {
        return { connId: p.connId, msg: JSON.parse(p.payload) as WireMessage };
    } catch {
        // A line that is not a wire message is a peer speaking a language this
        // build does not know. Dropping it beats tearing the link down.
        return null;
    }
}

async function defaultBridge(): Promise<TauriBridge> {
    // Dynamic so the browser bundle never pulls in Tauri's JS API, matching
    // how every other desktop touchpoint is gated.
    const core = await import("@tauri-apps/api/core");
    const event = await import("@tauri-apps/api/event");
    return {
        invoke: (cmd, args) => core.invoke(cmd, args),
        listen: async (name, cb) => {
            const un = await event.listen(name, (e) => cb(e.payload));
            return () => un();
        },
    };
}

interface Held {
    conn: PeerConn;
    kind: ConnectionKind;
    onMessage: ((m: WireMessage) => void)[];
    onClose: (() => void)[];
    open: boolean;
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
            conn: {
                id: remote,
                connectionType: () => entry.kind,
                send(msg) {
                    if (!entry.open) return;
                    // The shell refuses a send for one reason: it is not
                    // holding this connection. A peer that quit and an endpoint
                    // that stopped both land here, neither is retryable, and a
                    // peer going away is ordinary. So the link is dropped
                    // rather than left claiming to be up, and nothing about it
                    // reaches the debater as an error.
                    void shell
                        .invoke("collab_send", { connId, payload: JSON.stringify(msg) })
                        .catch(() => dropConn(connId));
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
                    // Already gone on the shell's side is the ordinary way a
                    // close races a peer that hung up first.
                    void shell.invoke("collab_close", { connId }).catch(() => {});
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
            for (const un of unlisten) un();
            unlisten.length = 0;
            onPeer = null;
            held.clear();
            await shell.invoke("collab_stop", {});
        },
    };
}
