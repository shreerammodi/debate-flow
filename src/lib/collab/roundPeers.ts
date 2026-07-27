/**
 * The peers a round has been shared with.
 *
 * Held for the open round only, and written into the sidecar on every save, so
 * that opening the file tomorrow re-dials the same partners with no ticket and
 * no interaction. A peer is only ever added: a partner who is offline right now
 * is still this round's partner, and forgetting them would cost a ticket to get
 * back.
 */

let heldRoundId: string | null = null;
let held: string[] = [];

/** Replaces the set, for a round that is being opened. */
export function setRoundPeers(roundId: string, peers: readonly string[]): void {
    heldRoundId = roundId;
    held = [...new Set(peers)];
}

/** Adds peers to the set, for a round that is already open. */
export function rememberRoundPeers(roundId: string, peers: readonly string[]): void {
    if (heldRoundId !== roundId) {
        setRoundPeers(roundId, peers);
        return;
    }
    for (const peer of peers) if (!held.includes(peer)) held.push(peer);
}

/** Empty for any round but the one being tracked, which is what a fresh open is. */
export function knownRoundPeers(roundId: string): string[] {
    return heldRoundId === roundId ? [...held] : [];
}

export function forgetRoundPeers(): void {
    heldRoundId = null;
    held = [];
}
