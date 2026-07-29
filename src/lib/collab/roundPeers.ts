/**
 * The peers a round has been shared with, and which of them read and do not
 * write.
 *
 * Held for the open round only, and written into the sidecar on every save, so
 * that opening the file tomorrow re-dials the same partners with no ticket and
 * no interaction. Remembering only ever adds: a partner who is offline right
 * now is still this round's partner, and forgetting them would cost a ticket to
 * get back. Cutting a peer loose is the one exception, because a debater who
 * does that means it to outlast the session.
 *
 * The read-only mark rides with the membership rather than in the contact table
 * alone. Membership with no grade beside it reads as the wider role, so a grant
 * kept only where a toast put it is destroyed by the gesture that most looks
 * like withdrawing trust.
 */

import type { Role } from "./types";

let heldRoundId: string | null = null;
let held: string[] = [];
/** The peers of this round that were admitted read-only. */
let readOnly: string[] = [];

/** Replaces the set, for a round that is being opened. */
export function setRoundPeers(
    roundId: string,
    peers: readonly string[],
    readOnlyPeers: readonly string[] = [],
): void {
    heldRoundId = roundId;
    held = [...new Set(peers)];
    readOnly = [...new Set(readOnlyPeers)];
}

/** Adds peers to the set, for a round that is already open. */
export function rememberRoundPeers(roundId: string, peers: readonly string[]): void {
    if (heldRoundId !== roundId) {
        setRoundPeers(roundId, peers);
        return;
    }
    for (const peer of peers) if (!held.includes(peer)) held.push(peer);
}

/**
 * Records a peer and what it was admitted as, so the grant outlives both the
 * session that made it and the contact table that graded it. A peer admitted
 * wider loses the mark: a later ticket is the debater deciding again.
 *
 * A grade is a membership, so this remembers the peer too. The two lists have
 * no way to disagree about who belongs.
 */
export function rememberRoundRole(roundId: string, peer: string, role: Role): void {
    if (heldRoundId !== roundId) return;
    if (!held.includes(peer)) held.push(peer);
    const marked = readOnly.includes(peer);
    if (role === "coach") {
        if (!marked) readOnly.push(peer);
    } else if (marked) {
        readOnly = readOnly.filter((p) => p !== peer);
    }
}

/** Empty for any round but the one being tracked, which is what a fresh open is. */
export function knownRoundPeers(roundId: string): string[] {
    return heldRoundId === roundId ? [...held] : [];
}

/** Of those peers, the ones a session has to keep read-only. */
export function knownRoundCoaches(roundId: string): string[] {
    return heldRoundId === roundId ? [...readOnly] : [];
}

/**
 * Drops one peer for good, for a debater who cut them loose. The set is
 * otherwise append-only, so without this the cut lasts only until the next open
 * re-dials them off the sidecar and admits them on membership alone.
 */
export function forgetRoundPeer(roundId: string, peer: string): void {
    if (heldRoundId !== roundId) return;
    held = held.filter((p) => p !== peer);
    readOnly = readOnly.filter((p) => p !== peer);
}

/** Drops the set, for a round that is being closed. Nothing is open to remember. */
export function forgetRoundPeers(): void {
    heldRoundId = null;
    held = [];
    readOnly = [];
}
