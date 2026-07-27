/**
 * A shared session: who is connected, what they may do, and keeping them in
 * step.
 *
 * The opt-in gate is here and nowhere else. With the master switch off this
 * function returns null and never touches the link factory, so no endpoint is
 * bound, no peer is dialled, no discovery record is published, and no relay is
 * contacted. Later phases may grow the body; they must not grow a second entry
 * point around it.
 *
 * The topology is a star and the host is the hub, so the host is the only side
 * that decides admission and the only side that has to enforce a read-only
 * role.
 */

import { setLocks } from "@/lib/grid/lockBridge";

import { collabSettings, type CollabSettings } from "./enabled";
import { admit, helloFrom, type HostPolicy } from "./handshake";
import type { DroppedCell } from "./merge";
import type { PeerConn, PeerLink, PeerLinkFactory, WireMessage } from "./peerLink";
import { claim, HEARTBEAT_MS, releaseCell, releasePeer, type Lock } from "./presence";
import { retryForever } from "./reconnect";
import { attachSync, type PeerSync } from "./sync";
import { mintTicket, parseTicket, type Ticket } from "./ticket";
import type { CollabDoc, Role } from "./types";

export interface CollabPeer {
    endpointId: string;
    role: Role;
    connectionType: "direct" | "relayed";
}

export interface CollabSession {
    endpointId: string;
    roundId: string;
    peers(): CollabPeer[];
    /** Mints the ticket the next peer presents. Replaces any unspent one. */
    share(role: Role): Ticket;
    /** Tells every peer about a local edit. */
    notifyLocalChange(): void;
    /**
     * Claims the cell this side has an editor open on, or releases it with
     * null. Sent immediately rather than on a tick, and refreshed on a
     * heartbeat so a frozen process stops holding it.
     */
    setPresence(cell: { sheetId: string; col: number; row: number } | null): void;
    stop(): Promise<void>;
}

export interface CollabSessionDeps {
    createLink: PeerLinkFactory;
    roundId: string;
    appVersion: string;
    doc(): CollabDoc;
    apply(incoming: CollabDoc): DroppedCell[];
    /** Injectable so a caller can drive the switch without a store write. */
    settings?: () => CollabSettings;
    /** Peers this round already knows, re-dialled silently when it opens. */
    dial?: string[];
    /** A ticket to present on the first dial, from a paste. */
    ticket?: string;
    /** What this side asks to be. The host's ticket is the authority. */
    role?: Role;
    onPeersChanged?: (peers: CollabPeer[]) => void;
    schedule?: (fn: () => void, ms: number) => () => void;
}

interface Live {
    conn: PeerConn;
    sync: PeerSync;
    peer: CollabPeer;
}

export async function startCollabSession(deps: CollabSessionDeps): Promise<CollabSession | null> {
    const settings = (deps.settings ?? collabSettings)();
    if (!settings.enabled) return null;

    const link: PeerLink = await deps.createLink({
        // mDNS reaches the machine across the room with no internet at all.
        // DNS discovery would publish this install to a public registry for a
        // session that is always invited by hand, so it is never an option.
        discovery: "mdns",
        relay: settings.relay,
    });
    const endpointId = await link.endpointId();

    const policy: HostPolicy = {
        roundId: deps.roundId,
        appVersion: deps.appVersion,
        pendingSecret: null,
        knownPeers: [...(deps.dial ?? [])],
    };
    const live = new Map<string, Live>();
    /** Cells peers have an editor open on. Advisory, and always expiring. */
    let locks: Lock[] = [];
    let stopped = false;

    function publishLocks(): void {
        setLocks(locks);
    }

    /** The cell this side holds, refreshed until the editor closes. */
    let held: { sheetId: string; col: number; row: number } | null = null;
    let cancelHeartbeat: (() => void) | null = null;
    const schedule =
        deps.schedule ??
        ((fn, ms) => {
            const id = setTimeout(fn, ms);
            return () => clearTimeout(id);
        });

    function broadcastPresence(): void {
        for (const l of live.values()) l.conn.send({ type: "presence", cell: held });
    }

    function armHeartbeat(): void {
        cancelHeartbeat = schedule(() => {
            if (stopped || !held) return;
            broadcastPresence();
            armHeartbeat();
        }, HEARTBEAT_MS);
    }

    function announce(): void {
        deps.onPeersChanged?.([...live.values()].map((l) => l.peer));
    }

    function onPresence(endpointId: string, cell: Lock | null): void {
        locks = cell ? claim(locks, cell) : releaseCell(locks, endpointId);
        publishLocks();
    }

    function track(conn: PeerConn, peer: CollabPeer, readOnly: boolean): PeerSync {
        // A peer's open editor claims a cell so this side sees it before
        // typing into it; the claim goes the moment the link does.
        conn.onMessage((msg) => {
            if (msg.type !== "presence") return;
            onPresence(
                peer.endpointId,
                msg.cell ? { endpointId: peer.endpointId, ...msg.cell, heldAt: Date.now() } : null,
            );
        });
        const sync = attachSync({
            conn,
            doc: deps.doc,
            apply: deps.apply,
            readOnly,
            endpointId,
            schedule: deps.schedule,
        });
        live.set(peer.endpointId, { conn, sync, peer });
        conn.onClose(() => {
            const held = live.get(peer.endpointId);
            if (held?.conn !== conn) return;
            held.sync.stop();
            live.delete(peer.endpointId);
            // A peer that is gone holds nothing, instantly and without waiting
            // for the heartbeat to lapse.
            locks = releasePeer(locks, peer.endpointId);
            publishLocks();
            announce();
        });
        announce();
        return sync;
    }

    // --- Host side ---------------------------------------------------------

    await link.listen((conn) => {
        if (stopped) return;
        // The first message must be a hello, and this handler answers exactly
        // one. Without disarming, the first delta that arrived afterwards
        // would read as a malformed hello and close a healthy link.
        let greeted = false;
        conn.onMessage((msg) => {
            if (greeted) return;
            greeted = true;

            const verdict = admit(msg, policy);
            if (!verdict.ok) {
                conn.send({ type: "helloAck", ok: false, reason: verdict.reason });
                conn.close();
                return;
            }
            if (msg.type !== "hello") return;
            if (verdict.spendSecret) policy.pendingSecret = null;
            if (!policy.knownPeers.includes(msg.endpointId)) {
                policy.knownPeers.push(msg.endpointId);
            }
            conn.send({ type: "helloAck", ok: true });
            const sync = track(
                conn,
                {
                    endpointId: msg.endpointId,
                    role: verdict.role,
                    connectionType: conn.connectionType(),
                },
                verdict.role === "coach",
            );
            // A guest may hold no file at all, so the host opens with the
            // whole document rather than a delta against a seed it cannot
            // assume.
            sync.sendState();
        });
    });

    // --- Guest side --------------------------------------------------------

    async function dialPeer(target: string, ticket?: string): Promise<void> {
        if (stopped) return;
        const conn = await link.dial(target, ticket);
        const secret = ticket ? (parseTicket(ticket)?.secret ?? undefined) : undefined;

        return new Promise<void>((resolve, reject) => {
            let answered = false;
            conn.onMessage((msg) => {
                if (answered || msg.type !== "helloAck") return;
                answered = true;
                if (!msg.ok) {
                    conn.close();
                    reject(new Error(msg.reason));
                    return;
                }
                track(
                    conn,
                    {
                        endpointId: target,
                        role: deps.role ?? "partner",
                        connectionType: conn.connectionType(),
                    },
                    false,
                );
                resolve();
            });
            conn.onClose(() => {
                if (!answered) reject(new Error("closed"));
            });

            // Sent last. A transport can answer synchronously, so the listener
            // has to be in place before the question goes out or the reply
            // lands on nobody.
            conn.send(
                helloFrom({
                    endpointId,
                    roundId: deps.roundId,
                    role: deps.role ?? "partner",
                    appVersion: deps.appVersion,
                    ticket: secret,
                }),
            );
        });
    }

    for (const target of deps.dial ?? []) {
        try {
            await dialPeer(target, deps.ticket);
        } catch {
            // A peer that is not up yet is ordinary, and a refusal is not a
            // reason to fail opening the round. Retry runs on its own.
            if (!stopped && deps.schedule) {
                retryForever({
                    dial: () => dialPeer(target),
                    schedule: deps.schedule,
                });
            }
        }
    }

    return {
        endpointId,
        roundId: deps.roundId,

        peers() {
            return [...live.values()].map((l) => l.peer);
        },

        share(role) {
            const ticket = mintTicket({
                endpointId,
                roundId: deps.roundId,
                role,
                relay: settings.relay,
            });
            policy.pendingSecret = ticket.secret;
            return ticket;
        },

        notifyLocalChange() {
            for (const l of live.values()) l.sync.notifyLocalChange();
        },

        setPresence(cell) {
            held = cell;
            broadcastPresence();
            cancelHeartbeat?.();
            cancelHeartbeat = null;
            // Only an open editor needs refreshing; a release is final.
            if (cell) armHeartbeat();
        },

        async stop() {
            stopped = true;
            for (const l of [...live.values()]) {
                l.sync.stop();
                l.conn.send({ type: "bye" });
                l.conn.close();
            }
            live.clear();
            cancelHeartbeat?.();
            cancelHeartbeat = null;
            held = null;
            locks = [];
            publishLocks();
            announce();
            await link.stop();
        },
    };
}
