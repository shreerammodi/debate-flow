/**
 * The one place a reason for decision changes hands.
 *
 * Each peer sends only its own notes. On the way out the local `rfd` register
 * is renamed to this peer's note path; on the way in, a peer's local `rfd` is
 * dropped. So there is exactly one writer per path and a merge cannot pit two
 * people's reasoning against each other: no text CRDT, no line diff, no
 * tombstone.
 *
 * The result is asymmetric on purpose. On your disk `rfd` is yours and their
 * reasoning sits under their EndpointId; on theirs it is the other way round.
 * That is what keeps `rfd` meaning the same thing it always meant, so the
 * exporters, the print view, the search index, and older builds all keep
 * working and the file stays at version 3.
 */

import type { CollabDoc } from "./types";

/** This machine owner's own notes. Never sent, never accepted. */
export const LOCAL_RFD_PATH = "scouting.decision.rfd";

/** Where one peer's notes live in everybody else's document. */
export function peerNotePath(endpointId: string): string {
    return `scouting.decision.peerNotes.${endpointId}`;
}

/** The document as this peer should send it: my notes, under my name. */
export function outgoingDoc(doc: CollabDoc, myEndpointId: string): CollabDoc {
    const mine = doc.round[LOCAL_RFD_PATH];
    if (!mine) return doc;
    const { [LOCAL_RFD_PATH]: _local, ...rest } = doc.round;
    return { ...doc, round: { ...rest, [peerNotePath(myEndpointId)]: mine } };
}

/**
 * This machine's own reasoning has one home, `rfd`. A copy of it under this
 * machine's own EndpointId is a peer's bookkeeping that came back, and it has
 * no business in the document it came from.
 */
export function dropSelfNote(doc: CollabDoc, myEndpointId: string): CollabDoc {
    const mine = peerNotePath(myEndpointId);
    if (!(mine in doc.round)) return doc;
    const { [mine]: _echo, ...rest } = doc.round;
    return { ...doc, round: rest };
}

/**
 * The document as it may be applied here, which refuses two things.
 *
 * A peer sending a local `rfd` is either an older build or a modified client;
 * either way those are their notes and they do not belong in this editor.
 *
 * And every peer holds this machine's own note under this machine's id,
 * because that is how it was sent to them. Echoed back it would land beside
 * the `rfd` it was written in, be shown to its own author as a partner's
 * reasoning, and never move again: the author edits `rfd`, and nothing on this
 * machine writes the copy.
 */
export function incomingDoc(doc: CollabDoc, myEndpointId: string): CollabDoc {
    const theirs = dropSelfNote(doc, myEndpointId);
    if (!(LOCAL_RFD_PATH in theirs.round)) return theirs;
    const { [LOCAL_RFD_PATH]: _local, ...rest } = theirs.round;
    return { ...theirs, round: rest };
}
