/**
 * Turns the presence table into something the grid can paint.
 *
 * A cell a peer is on is marked so the debater can see where their partner is
 * working without asking, and a cell a peer is editing is marked harder,
 * before the debater tries to type into it, so a refusal is predictable rather
 * than a surprise. Liveness is not this module's rule to invent: `expire` in
 * presence owns the TTL, and an entry past it decorates nothing, exactly as it
 * holds nothing.
 */

import { expire, lockAt, presenceAt, type Presence } from "@/lib/collab/presence";

import type { ModelCol } from "./colSpace";

/** Marks a cell a peer's cursor is on. */
export const PEER_CLASS = "ebb-peer";
/** Marks a cell a peer has an editor open on, which also wears PEER_CLASS. */
export const LOCK_CLASS = "ebb-locked";

/** The peer on this cell, or null when nobody is. */
export function presenceOn(
    list: readonly Presence[],
    sheetId: string,
    col: ModelCol,
    row: number,
    now: number,
): Presence | null {
    return presenceAt(expire(list, now), sheetId, col, row);
}

/**
 * The one character the corner badge shows. A name is what the debater reads
 * everywhere else, so its first letter is what identifies the partner here;
 * a name that is only punctuation still has to leave a mark, so the fallback
 * is a bullet rather than an empty badge.
 */
export function peerInitial(name: string): string {
    const letter = [...name].find((c) => /[\p{L}\p{N}]/u.test(c));
    return letter ? letter.toUpperCase() : "*";
}

/**
 * Who has an editor open on that cell, named for a hint or tooltip. `nameOf`
 * resolves an endpoint id to a display name, which is the contacts list's job,
 * not this module's.
 */
export function lockLabel(
    list: readonly Presence[],
    sheetId: string,
    col: ModelCol,
    row: number,
    now: number,
    nameOf: (endpointId: string) => string,
): string | null {
    const held = lockAt(expire(list, now), sheetId, col, row);
    return held ? nameOf(held.endpointId) : null;
}
