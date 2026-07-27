/**
 * The replica across a restart.
 *
 * Recovery is best-effort by construction: a sidecar is only ever an
 * accelerator, so a missing, stale, malformed, or unreadable one falls back to
 * seeding from the file, which two peers do identically. Persisting is
 * likewise never allowed to fail a save; the flow itself is already on disk by
 * the time this runs.
 *
 * The replica is maintained whether or not a peer exists, so recovery always
 * seeds. Only the disk write is gated on the master switch: a debater who has
 * never turned shared editing on gets no extra file.
 */

import type { FlowRound } from "@/lib/model/flow";

import { collabSettings } from "./enabled";
import { hashText } from "./hash";
import { getReplica, healReplica, replicaRoundId, seedReplica } from "./replica";
import { knownRoundPeers, rememberRoundPeers } from "./roundPeers";
import { parseSidecar, serializeSidecar } from "./sidecar";
import { getSidecarFs } from "./sidecarFs";

/**
 * Seeds the replica for a round that is being opened, and answers with the
 * peers that round remembers, which is who a session re-dials.
 */
export async function recoverReplica(round: FlowRound, flowText: string): Promise<string[]> {
    if (!collabSettings().enabled) {
        seedReplica(round);
        return [];
    }
    let recovered = null;
    try {
        const fs = await getSidecarFs();
        recovered = parseSidecar(await fs.read(round.id), round.id, hashText(flowText));
    } catch {
        // A broken config directory is not a reason to refuse to open a round.
    }
    seedReplica(round, "", recovered?.doc ?? null);
    // Remembered rather than replaced: a round taken from an invitation knows
    // its host before any sidecar for it exists.
    rememberRoundPeers(round.id, recovered?.peers ?? []);
    return knownRoundPeers(round.id);
}

/** Repairs any drift, then makes the replica durable beside the saved file. */
export async function persistReplica(round: FlowRound, flowText: string): Promise<void> {
    if (!collabSettings().enabled) return;
    if (replicaRoundId() !== round.id) return;
    healReplica(round);
    const doc = getReplica();
    if (!doc) return;
    try {
        const fs = await getSidecarFs();
        await fs.write(
            round.id,
            serializeSidecar({
                roundId: round.id,
                flowHash: hashText(flowText),
                peers: knownRoundPeers(round.id),
                doc,
            }),
        );
    } catch {
        // The flow itself is already saved. A sidecar that did not land only
        // costs a re-seed on the next open.
    }
}
