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
import { projectDoc } from "./doc";
import { collabSettings } from "./enabled";
import { announceInvite } from "./inbox";
import type { InviteNotice } from "./invite";
import { startInviteListener, type InviteListener } from "./inviteListener";
import { lossMessage } from "./lossReport";
import { broadcastName } from "./machineName";
import { merge, type DroppedCell } from "./merge";
import { createPeerLinkFor } from "./peerLink";
import {
    adoptReplicaActor,
    getReplica,
    replaceReplicaDoc,
    replicaActor,
    replicaRoundId,
    seedReplica,
    setLocalChangeListener,
} from "./replica";
import { dropSelfNote } from "./rfdSync";
import { knownRoundPeers, rememberRoundPeers } from "./roundPeers";
import { startCollabSession, type CollabPeer, type CollabSession } from "./session";
import type { CollabDoc } from "./types";

let session: CollabSession | null = null;
/** Bound between rounds so a saved contact's invite has somewhere to land. */
let listener: InviteListener | null = null;
/** In-flight listener change, so two callers cannot bind two endpoints. */
let watching: Promise<void> | null = null;
/**
 * A session is coming up but has not been assigned yet. There is one endpoint
 * per install: a listener that bound during this window would share it with
 * the session, hear the session's own peers as diallers, and hang up on them.
 */
let starting = false;
/** The name each peer was last offered under, so one session asks once. */
const offered = new Map<string, string>();

/**
 * A peer nobody has saved is worth one offer, because the alternative is
 * trading keys by hand next time. The name defaults to the one they broadcast,
 * falling back to the short id, and is theirs to change in Settings; what
 * matters is the id behind it.
 *
 * Asked again when a peer that greeted this machine namelessly has since said
 * what to call them. The offer carries the name it will save, so one made
 * before the name arrived would save a short EndpointId as this partner's name
 * for good. The toast is addressed per peer, so the later offer replaces the
 * earlier one in place rather than stacking a second question beside it.
 */
function offerToSave(peers: CollabPeer[]): void {
    const contacts = useFlowStore.getState().contacts;
    for (const peer of peers) {
        if (contacts[peer.endpointId]) continue;
        const name = contactName(contacts, peer.endpointId, peer.name);
        if (offered.get(peer.endpointId) === name) continue;
        offered.set(peer.endpointId, name);
        toast(`Save ${name} as a ${peer.role}?`, {
            id: `collab-save-${peer.endpointId}`,
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
        name: contactName(contacts, p.endpointId, p.name),
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
 * the replica, write the round it now describes, move the grid under the apply
 * rules, and tell the user about anything the merge buried.
 *
 * The store holds the projection of the replica, so this is what puts a
 * partner's text in front of the user, into the autosave, and into every
 * export. The grid follows separately because what may be painted over is
 * governed by the apply rules, which is a different question from what the
 * document says.
 */
export function applyRemoteDoc(round: FlowRound, incoming: CollabDoc): DroppedCell[] {
    const before = getReplica();
    if (!before) return [];
    const result = merge(before, incoming);
    replaceReplicaDoc(result.doc);
    const store = useFlowStore.getState();
    // The round on screen, not the one the session opened with: createdAt and
    // updatedAt belong to this file and no partner ever sends them.
    const base = store.round?.id === result.doc.roundId ? store.round : round;
    store.applyRemoteRound(projectDoc(result.doc, base));
    applyRemote(before, result.doc);
    // A delete leaves no mark on the grid, so this is the only place the user
    // learns their text is gone.
    const loss = lossMessage(store.contacts, result.dropped, replicaActor());
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
    // Claimed before the release, and held until the session is assigned, so
    // no listener can bind the endpoint in between.
    starting = true;
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
            roundId: round.id,
            // The filename is what a debater calls this round everywhere else,
            // so it is what an invite names.
            roundLabel: store.docPath ? basename(store.docPath).replace(/\.ebb$/i, "") : "",
            appVersion: await getCurrentVersion(),
            displayName: await broadcastName(),
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
        starting = false;
        useCollabStore.getState().reset();
        await syncInviteWatch();
        throw err;
    }
    starting = false;

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
    // An earlier session may have left this machine's own note in the file
    // under this machine's id, where the drawer reads it as a partner's. This
    // is the first moment the id is known, so it is where that is undone.
    const held = getReplica();
    const clean = held && dropSelfNote(held, session.endpointId);
    if (clean && clean !== held) {
        replaceReplicaDoc(clean);
        useFlowStore.getState().applyRemoteRound(projectDoc(clean, round));
    }
    // Push, not poll: every op the grid and the store record is offered to the
    // peers the moment it lands, coalesced a frame later by the sync.
    setLocalChangeListener(notifyLocalChange);
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

/**
 * Ends the session, whatever the transport thinks. The state above the link is
 * already torn down by the time the link is asked to stop, so a shell that
 * refuses cannot leave a half-ended session behind - and End session is a
 * button a debater presses mid-round, which must never answer with an error.
 */
export async function endSession(): Promise<void> {
    const held = session;
    session = null;
    setLocalChangeListener(null);
    offered.clear();
    try {
        await held?.stop();
    } catch {
        // The endpoint is going away with the session either way.
    }
    // After the stop, never before it: a session announces an empty peer list
    // on its way out, and a reset that ran first would be overwritten by it
    // and leave the chip saying "connecting" for a session that is over.
    useCollabStore.getState().reset();
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
    const live = session;
    const held = live ?? (await startForRound(round, [endpointId]));
    if (!held) throw new Error("Turn on shared editing in Settings first");
    rememberRoundPeers(round.id, [endpointId]);
    // Opening a session for this round dials the contact on the way up, and
    // that dial is the invitation. Dialling again would put a second notice on
    // their screen for one share.
    if (!live) return;
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
            const wanted = collabSettings().enabled && !session && !starting;
            if (wanted !== (listener !== null)) {
                if (!wanted) {
                    await dropListener();
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

/**
 * Lets go of the listener. For callers already inside the watch chain.
 *
 * A shell that refuses to stop is not a reason to keep a handle nobody can
 * use: the endpoint is dropped here either way, so a session asking for it
 * next is never blocked by a failed release.
 */
async function dropListener(): Promise<void> {
    const held = listener;
    listener = null;
    try {
        await held?.stop();
    } catch {
        // Nothing above this holds the listener any more.
    }
}

/**
 * Lets go of the listener from outside the watch chain, queued behind a bind
 * already in flight. Releasing ahead of that bind would let it land afterwards
 * and take the endpoint a session is on its way to needing.
 */
async function releaseInviteWatch(): Promise<void> {
    const next = (watching ?? Promise.resolve()).catch(() => {}).then(dropListener);
    watching = next;
    await next;
}
