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

import { toast } from "sonner";

import { applyRemote } from "@/lib/grid/remoteBridge";
import type { FlowRound } from "@/lib/model/flow";
import { basename } from "@/lib/persistence/flowPaths";
import { useCollabStore, type CollabPeerView } from "@/lib/store/useCollabStore";
import { useFlowStore } from "@/lib/store/useFlowStore";
import { getCurrentVersion } from "@/lib/update/adapter";

import { addContact, contactName, type Contact } from "./contacts";
import { collabSettings } from "./enabled";
import { announceInvite } from "./inbox";
import type { InviteNotice } from "./invite";
import { startInviteListener, type InviteListener } from "./inviteListener";
import { lossMessage } from "./lossReport";
import { merge, type DroppedCell } from "./merge";
import { createPeerLinkFor } from "./peerLink";
import {
    adoptReplicaActor,
    getReplica,
    replicaActor,
    replicaRoundId,
    seedReplica,
} from "./replica";
import { knownRoundPeers, rememberRoundPeers } from "./roundPeers";
import { startCollabSession, type CollabPeer, type CollabSession } from "./session";
import { createShadow } from "./shadow";
import type { CollabDoc } from "./types";

let session: CollabSession | null = null;
/** Bound between rounds so a saved contact's invite has somewhere to land. */
let listener: InviteListener | null = null;
/** In-flight listener change, so two callers cannot bind two endpoints. */
let watching: Promise<void> | null = null;
/** Peers already offered as a contact, so one session asks about each once. */
const offered = new Set<string>();

/** An EndpointId is a long key; a chip shows the first eight characters. */
function shortName(endpointId: string): string {
    return endpointId.slice(0, 8);
}

/**
 * A peer nobody has saved is worth one offer, because the alternative is
 * trading keys by hand next time. The name defaults to the short id
 * and is theirs to change in Settings; what matters is the id behind it.
 */
function offerToSave(peers: CollabPeer[]): void {
    const contacts = useFlowStore.getState().contacts;
    for (const peer of peers) {
        if (contacts[peer.endpointId] || offered.has(peer.endpointId)) continue;
        offered.add(peer.endpointId);
        const name = shortName(peer.endpointId);
        toast(`Save ${name} as a ${peer.role}?`, {
            duration: 20_000,
            action: {
                label: "Save",
                onClick: () => saveContact(peer.endpointId, { name, role: peer.role }),
            },
        });
    }
}

function publish(peers: CollabPeer[]): void {
    const contacts = useFlowStore.getState().contacts;
    const view: CollabPeerView[] = peers.map((p) => ({
        endpointId: p.endpointId,
        name: contactName(contacts, p.endpointId),
        role: p.role,
        connectionType: p.connectionType,
    }));
    useCollabStore.getState().setPeers(view);
    useCollabStore.getState().setStatus(view.length > 0 ? "connected" : "connecting");
    if (session)
        rememberRoundPeers(
            session.roundId,
            view.map((p) => p.endpointId),
        );
    offerToSave(peers);
}

/**
 * Everything that happens to a remote document on the way in: merge it, keep
 * the replica, move the grid under the apply rules, and tell the user about
 * anything the merge buried.
 *
 * The replica is correct the instant the merge lands. What reaches the grid is
 * governed by the apply rules, which is a separate question from what the
 * document says.
 */
export function applyRemoteDoc(round: FlowRound, incoming: CollabDoc): DroppedCell[] {
    const before = getReplica();
    if (!before) return [];
    const result = merge(before, incoming);
    seedReplica(round, replicaActor(), result.doc);
    applyRemote(before, result.doc);
    // A delete leaves no mark on the grid, so this is the only place the user
    // learns their text is gone.
    const loss = lossMessage(useFlowStore.getState().contacts, result.dropped, replicaActor());
    if (loss) toast.warning(loss, { duration: 10_000 });
    return result.dropped;
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
    if (session?.roundId === round.id) return session;
    // A session speaks for one round, so opening another one ends it.
    if (session) await endSession();
    // One endpoint per install, so the idle listener lets go of it first.
    await releaseInviteWatch();

    useCollabStore.getState().setStatus("connecting");
    // The replica is already live from opening the round; the session reads it
    // rather than holding a second copy.
    if (replicaRoundId() !== round.id) seedReplica(round);

    const store = useFlowStore.getState();
    try {
        session = await startCollabSession({
            createLink: createPeerLinkFor,
            // Shadow mode is read once, at session start. Flipping it mid-round
            // would leave the two sides disagreeing about what has been applied.
            shadow: store.shadowMode
                ? createShadow({ doc: () => getReplica() as CollabDoc, base: () => round })
                : undefined,
            onShadow: (entry) => useCollabStore.getState().pushShadow(entry),
            roundId: round.id,
            // The filename is what a debater calls this round everywhere else,
            // so it is what an invite names.
            roundLabel: store.docPath ? basename(store.docPath).replace(/\.ebb$/i, "") : "",
            appVersion: await getCurrentVersion(),
            doc: () => getReplica() as CollabDoc,
            apply: (incoming) => applyRemoteDoc(round, incoming),
            dial: knownPeers,
            onPeersChanged: publish,
            contacts: () => useFlowStore.getState().contacts,
            onInvite: announceInvite,
        });
    } catch (err) {
        // A chip left saying "connecting" would outlast the corner message and
        // read as a session that is still coming up.
        session = null;
        useCollabStore.getState().reset();
        await syncInviteWatch();
        throw err;
    }

    if (!session) {
        useCollabStore.getState().reset();
        await syncInviteWatch();
        return session;
    }
    rememberRoundPeers(round.id, knownPeers);
    useCollabStore.getState().setEndpointId(session.endpointId);
    // Every cell written from here carries this machine's own identity, so a
    // cell it inserts can never collide with one a peer inserts at the same
    // position.
    adoptReplicaActor(session.endpointId);
    return session;
}

/**
 * Re-dials the peers a round remembers, which is what makes a reconnect cost
 * no ticket and no interaction. A round nobody has shared stays offline.
 */
export async function resumeSession(round: FlowRound): Promise<CollabSession | null> {
    const peers = knownRoundPeers(round.id);
    if (peers.length === 0) return null;
    return startForRound(round, peers);
}

/** Tells the live session an edit landed. A no-op with no session. */
export function notifyLocalChange(): void {
    session?.notifyLocalChange();
}

export async function endSession(): Promise<void> {
    const held = session;
    session = null;
    offered.clear();
    useCollabStore.getState().reset();
    await held?.stop();
    await syncInviteWatch();
}

export async function disconnectPeer(endpointId: string): Promise<void> {
    // One peer leaving is not the session ending, so the rest stay up.
    const held = session;
    if (!held) return;
    const peers = held.peers().filter((p) => p.endpointId !== endpointId);
    publish(peers);
}

/**
 * Dials a contact directly, with no ticket: their EndpointId already
 * authorizes, which is what a contact is for.
 */
export async function inviteContact(round: FlowRound, endpointId: string): Promise<void> {
    const held = session ?? (await startForRound(round, [endpointId]));
    if (!held) throw new Error("Turn on shared editing in Settings first");
    rememberRoundPeers(round.id, [endpointId]);
    if (held.peers().some((p) => p.endpointId === endpointId)) return;
    // A session was already up, so the contact is dialled onto it.
    await held.invite(endpointId);
}

/** Saves a peer so the next round needs no ticket. Only ever from one click. */
export function saveContact(endpointId: string, contact: Contact): void {
    const store = useFlowStore.getState();
    store.setContacts(addContact(store.contacts, endpointId, contact));
}

/**
 * Binds or releases the idle listener to match the world: it is up exactly
 * when shared editing is on and no session is holding the endpoint. Called on
 * boot, when the master switch moves, and at both ends of a session.
 */
export async function syncInviteWatch(): Promise<void> {
    // Binding an endpoint takes a moment, so callers are serialized rather
    // than allowed to race a listener nobody holds a handle to.
    const next = (watching ?? Promise.resolve())
        .catch(() => {})
        .then(async () => {
            const wanted = collabSettings().enabled && !session;
            if (wanted !== (listener !== null)) {
                if (!wanted) {
                    await releaseInviteWatch();
                } else {
                    try {
                        listener = await startInviteListener({
                            createLink: createPeerLinkFor,
                            contacts: () => useFlowStore.getState().contacts,
                            onInvite: (notice: InviteNotice) => announceInvite(notice),
                        });
                    } catch {
                        // An endpoint that will not bind costs an invitation, and
                        // nothing else. Ending a session, or opening the app, still
                        // succeeds.
                        listener = null;
                    }
                }
            }
            // The identity outlives any one binding, so it is published whenever
            // one is in hand rather than only at the moment it binds.
            if (listener) useCollabStore.getState().setEndpointId(listener.endpointId);
        });
    watching = next;
    await next;
}

async function releaseInviteWatch(): Promise<void> {
    const held = listener;
    listener = null;
    await held?.stop();
}
