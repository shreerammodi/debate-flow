/**
 * Turns the advisory lock table into something the grid can paint.
 *
 * A locked cell is marked before the debater tries to type, so a refusal is
 * predictable rather than surprising. Liveness is not this module's rule to
 * invent: `expire` in presence owns the TTL, and a lock past it decorates
 * nothing, exactly as it holds nothing.
 */

import { expire, lockAt, type Lock } from "@/lib/collab/presence";

/** Marks a cell a peer currently holds. */
export const LOCK_CLASS = "ebb-locked";

/** The decoration class for a held cell, or null when it is free. */
export function lockClassFor(
    locks: readonly Lock[],
    sheetId: string,
    col: number,
    row: number,
    now: number,
): string | null {
    return lockAt(expire(locks, now), sheetId, col, row) ? LOCK_CLASS : null;
}

/**
 * Who holds that cell, named for a hint or tooltip. `nameOf` resolves an
 * endpoint id to a display name, which is the contacts list's job, not this
 * module's.
 */
export function lockLabel(
    locks: readonly Lock[],
    sheetId: string,
    col: number,
    row: number,
    now: number,
    nameOf: (endpointId: string) => string,
): string | null {
    const held = lockAt(expire(locks, now), sheetId, col, row);
    return held ? nameOf(held.endpointId) : null;
}
