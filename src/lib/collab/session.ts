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

import type { Contacts } from "./contacts";
import { collabSettings, type CollabSettings } from "./enabled";
import { admit, helloFrom, type HostPolicy } from "./handshake";
import { INVITED, inviteFrom, type InviteNotice } from "./invite";
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
    /**
     * What the peer calls themselves, when they said. A suggestion only: a
     * saved contact's name is the receiver's own word for them and wins.
     */
    name?: string;
}

export interface CollabSession {
    endpointId: string;
    roundId: string;
    peers(): CollabPeer[];
    /** Mints the ticket the next peer presents. Replaces any unspent one. */
    share(role: Role): Ticket;
    /** Dials a peer this round already trusts, with no ticket. */
    invite(endpointId: string): Promise<void>;
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
    /** What this side calls the round, so an invite it sends can name it. */
    roundLabel?: string;
    /** What this side calls itself, carried to every peer it greets. */
    displayName?: string;
    /**
     * The contact table, consulted only to decide whether a dial this session
     * cannot admit is an invite worth showing. Absent means every refusal is
     * silent.
     */
    contacts?: () => Contacts;
    onInvite?: (notice: InviteNotice) => void;
}

interface Live {
    conn: PeerConn;
    sync: PeerSync;
    peer: CollabPeer;
}

/**
 * Whether a dial ended in the far side taking the invite rather than joining.
 * The notice landed, so the caller has nothing to report and nothing to retry.
 */
function isInvited(err: unknown): boolean {
    return err instanceof Error && err.message === INVITED;
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
        pending: null,
        knownPeers: [...(deps.dial ?? [])],
        roles: {},
    };
    const live = new Map<string, Live>();
    /** Peers this side reached out to, and so is the one to reach out again. */
    const dialled = new Set<string>();
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
            apply: (incoming) => {
                const dropped = deps.apply(incoming);
                // The star has the host for a hub, and a hub that does not
                // pass a change on is not one: two partners and a coach would
                // each see the host's typing at once and each other's only
                // when the repair tick came round, seconds later. A guest
                // holds one peer, the host, so this is the host's job alone
                // and stops after one hop.
                for (const [id, other] of live) {
                    if (id !== peer.endpointId) other.sync.notifyLocalChange();
                }
                return dropped;
            },
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
            // A link that blips mid-round is the ordinary case in a gym full
            // of laptops, and the dial that opened this one only retried while
            // the session was coming up. Without re-arming here, the first
            // drop is the last thing that ever happens to this peer.
            //
            // Only the side that dialled redials: the host cannot reach a
            // guest that has not spoken, and both sides trying would race two
            // connections into one slot.
            if (!stopped && dialled.has(peer.endpointId)) redial(peer.endpointId);
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

            // The connection's own id, not the one the hello claims: iroh
            // proved the far side holds that key, and everything downstream is
            // filed under it.
            const remoteId = conn.id;
            const verdict = admit(msg, policy, remoteId);
            if (!verdict.ok) {
                // A contact dialling about a round this side is not holding is
                // offering it, not failing to join one. The notice is all that
                // crosses; joining is the receiver's own move.
                const notice = deps.contacts
                    ? inviteFrom(msg, deps.contacts(), policy.roundId, remoteId)
                    : null;
                if (notice && deps.onInvite) {
                    deps.onInvite(notice);
                    conn.send({ type: "helloAck", ok: false, reason: INVITED });
                } else {
                    conn.send({ type: "helloAck", ok: false, reason: verdict.reason });
                }
                conn.close();
                return;
            }
            if (msg.type !== "hello") return;
            if (verdict.spendSecret) policy.pending = null;
            if (!policy.knownPeers.includes(remoteId)) policy.knownPeers.push(remoteId);
            policy.roles[remoteId] = verdict.role;
            conn.send({ type: "helloAck", ok: true, name: deps.displayName });
            const sync = track(
                conn,
                {
                    endpointId: remoteId,
                    role: verdict.role,
                    connectionType: conn.connectionType(),
                    name: typeof msg.name === "string" ? msg.name : undefined,
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
        dialled.add(target);
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
                        name: typeof msg.name === "string" ? msg.name : undefined,
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
                    label: deps.roundLabel,
                    name: deps.displayName,
                }),
            );
        });
    }

    /**
     * Keeps trying a peer until it answers.
     *
     * On the session's own scheduler, not the injectable one: a session that
     * only retried when a test handed it a clock would never reconnect on a
     * real machine, which is the only place a link actually drops.
     */
    function redial(target: string): void {
        retryForever({ dial: () => dialPeer(target), schedule });
    }

    for (const target of deps.dial ?? []) {
        try {
            await dialPeer(target, deps.ticket);
        } catch (err) {
            // A peer that is not up yet is ordinary, and a refusal is not a
            // reason to fail opening the round. Retry runs on its own, except
            // against a peer who answered: they heard this round offered and
            // are not holding it, so dialling again would only repeat itself.
            if (!stopped && !isInvited(err)) redial(target);
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
            // The role rides with the secret. What the ticket grants is the
            // host's to decide, so a coach's ticket cannot be spent as a
            // partner by a guest that simply says it is one.
            policy.pending = { secret: ticket.secret, role };
            return ticket;
        },

        async invite(target) {
            // A contact is admitted by EndpointId, so it joins the known list
            // before the dial rather than presenting a secret.
            if (!policy.knownPeers.includes(target)) policy.knownPeers.push(target);
            try {
                await dialPeer(target);
            } catch (err) {
                // They are on another round, so the dial delivered a notice
                // instead of a peer. That is the invite working.
                if (!isInvited(err)) throw err;
            }
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
