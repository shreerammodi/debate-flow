/**
 * The one live session, and the store the chip reads.
 *
 * Module state rather than store state, for the same reason the replica is:
 * the session is a connection, not a value a component renders. What the chip
 * needs - a status and a peer list - is pushed into `useCollabStore` as it
 * changes, and nothing else about the session is visible to React.
 *
 * Every route in and out of a session passes through here, so the master
 * switch is enforced in exactly one place.
 */

import type { FlowRound } from "@/lib/model/flow";
import { useCollabStore, type CollabPeerView } from "@/lib/store/useCollabStore";
import { getCurrentVersion } from "@/lib/update/adapter";

import { collabSettings } from "./enabled";
import { merge, type DroppedCell } from "./merge";
import { createPeerLinkFor } from "./peerLink";
import { getReplica, replicaRoundId, seedReplica } from "./replica";
import { startCollabSession, type CollabPeer, type CollabSession } from "./session";
import type { CollabDoc } from "./types";

let session: CollabSession | null = null;

/** An EndpointId is 52 characters of base32; a chip shows the first eight. */
function shortName(endpointId: string): string {
    return endpointId.slice(0, 8);
}

function publish(peers: CollabPeer[]): void {
    const view: CollabPeerView[] = peers.map((p) => ({
        endpointId: p.endpointId,
        name: shortName(p.endpointId),
        role: p.role,
        connectionType: p.connectionType,
    }));
    useCollabStore.getState().setPeers(view);
    useCollabStore.getState().setStatus(view.length > 0 ? "connected" : "connecting");
}

export function currentSession(): CollabSession | null {
    return session;
}

/**
 * Opens a session for the round already loaded. Returns null when shared
 * editing is off, which is the whole gate: nothing below this line runs.
 */
export async function startForRound(
    round: FlowRound,
    knownPeers: string[] = [],
): Promise<CollabSession | null> {
    if (!collabSettings().enabled) return null;
    if (session) return session;

    useCollabStore.getState().setStatus("connecting");
    // The replica is already live from opening the round; the session reads it
    // rather than holding a second copy.
    if (replicaRoundId() !== round.id) seedReplica(round);

    session = await startCollabSession({
        createLink: createPeerLinkFor,
        roundId: round.id,
        appVersion: await getCurrentVersion(),
        doc: () => getReplica() as CollabDoc,
        apply: (incoming): DroppedCell[] => {
            const held = getReplica();
            if (!held) return [];
            const result = merge(held, incoming);
            seedReplica(round, "", result.doc);
            return result.dropped;
        },
        dial: knownPeers,
        onPeersChanged: publish,
    });

    if (!session) useCollabStore.getState().reset();
    return session;
}

/** Tells the live session an edit landed. A no-op with no session. */
export function notifyLocalChange(): void {
    session?.notifyLocalChange();
}

export async function endSession(): Promise<void> {
    const held = session;
    session = null;
    useCollabStore.getState().reset();
    await held?.stop();
}

export async function disconnectPeer(endpointId: string): Promise<void> {
    // One peer leaving is not the session ending, so the rest stay up.
    const held = session;
    if (!held) return;
    const peers = held.peers().filter((p) => p.endpointId !== endpointId);
    publish(peers);
}
