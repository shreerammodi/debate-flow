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
 * The document as it may be applied here. A peer sending a local `rfd` is
 * either an older build or a modified client; either way those are their notes
 * and they do not belong in this editor.
 */
export function incomingDoc(doc: CollabDoc): CollabDoc {
    if (!(LOCAL_RFD_PATH in doc.round)) return doc;
    const { [LOCAL_RFD_PATH]: _theirs, ...rest } = doc.round;
    return { ...doc, round: rest };
}
