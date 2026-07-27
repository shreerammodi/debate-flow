/**
 * Who may join, and what they are allowed to do.
 *
 * Three rules carry the weight. A single-use secret admits the first peer that
 * presents it, and that peer's EndpointId admits it forever after, which is
 * what turns every later reconnect into no interaction at all. An unknown peer
 * with no valid secret is refused with no UI whatsoever: otherwise anyone who
 * learns your EndpointId can put notifications on your screen mid-round. And
 * every one of those decisions is made about the endpoint the transport
 * authenticated, never the one the hello claims to come from.
 */

import { PROTOCOL_MAJOR, type WireMessage } from "./peerLink";
import type { Role } from "./types";

export interface HostPolicy {
    /** The only round this host will talk about. */
    roundId: string;
    appVersion: string;
    /**
     * The unspent ticket, and what it grants. The role is the host's to give:
     * a coach's ticket admits a coach however the guest introduces itself.
     */
    pending: { secret: string; role: Role } | null;
    /** Peers already admitted once, which need no secret again. */
    knownPeers: string[];
    /** What an already-known peer was admitted as. Absent means partner. */
    roles: Record<string, Role>;
}

export type Admission =
    | { ok: true; role: Role; spendSecret: boolean }
    /** `silent` suppresses every surface, down to a chip flicker. */
    | { ok: false; reason: string; silent: boolean };

export function helloFrom(input: {
    endpointId: string;
    roundId: string;
    role: Role;
    appVersion: string;
    ticket?: string;
    /** What this side calls the round, so an invite can name it. */
    label?: string;
    /** What this side calls itself, so a peer has something to save. */
    name?: string;
}): WireMessage {
    const hello: Extract<WireMessage, { type: "hello" }> = {
        type: "hello",
        protocol: PROTOCOL_MAJOR,
        app: input.appVersion,
        endpointId: input.endpointId,
        roundId: input.roundId,
        role: input.role,
        // Shipped from day one so the first real skew can be negotiated
        // instead of refused.
        capabilities: [],
    };
    if (input.label) hello.label = input.label;
    if (input.name) hello.name = input.name;
    return input.ticket ? { ...hello, ticket: input.ticket } : hello;
}

/**
 * Compares without leaking how far the match got. A wrong guess should tell an
 * attacker nothing beyond "wrong", which is the standard the loopback bridge
 * is already held to.
 */
function secretMatches(a: string, b: string): boolean {
    if (a.length !== b.length) return false;
    let diff = 0;
    for (let i = 0; i < a.length; i++) diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
    return diff === 0;
}

const SILENT = { ok: false as const, reason: "refused", silent: true };

/**
 * `remoteId` is the endpoint the transport authenticated: iroh proved the far
 * side holds that key before a byte of this message existed. `msg.endpointId`
 * is a string the far side typed. They must agree, or the peer is claiming to
 * be somebody else and every later decision - who is admitted, whose name the
 * chip shows, whose reasoning a note is filed under - would be about the
 * wrong person.
 */
export function admit(msg: WireMessage, policy: HostPolicy, remoteId: string): Admission {
    if (msg.type !== "hello") return SILENT;
    if (msg.role !== "partner" && msg.role !== "coach") return SILENT;

    // A version skew is the one refusal a debater is told about, because the
    // fix is theirs to make and the corner can name both sides.
    if (msg.protocol !== PROTOCOL_MAJOR) {
        return {
            ok: false,
            reason: `they are on ebb ${msg.app}, you are on ebb ${policy.appVersion}`,
            silent: false,
        };
    }
    if (msg.roundId !== policy.roundId) return SILENT;
    if (msg.endpointId !== remoteId) return SILENT;

    if (policy.knownPeers.includes(remoteId)) {
        // A peer this round already knew but never graded is a partner: they
        // were admitted to it before, and the round remembers only that.
        return { ok: true, role: policy.roles[remoteId] ?? "partner", spendSecret: false };
    }
    if (policy.pending && msg.ticket && secretMatches(msg.ticket, policy.pending.secret)) {
        // The role the host granted, never the one the guest asked for.
        return { ok: true, role: policy.pending.role, spendSecret: true };
    }
    return SILENT;
}
