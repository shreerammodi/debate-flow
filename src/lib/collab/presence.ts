/**
 * Advisory cell locks.
 *
 * Opening an editor claims the cell so the other side sees it before they
 * start typing, which makes a refusal predictable rather than surprising. The
 * claim is advisory: there is no coordinator, so two peers can hold one cell
 * during a partition and last-writer-wins settles it underneath.
 *
 * What matters more than the lock is that it always goes away. A cell that
 * stays locked because a peer vanished is worse than no locking at all, so
 * there are three releases and only the last is timed: the editor closing, the
 * connection dropping, and a TTL for a frozen process on a live link.
 */

/** Refresh cadence while an editor stays open. */
export const HEARTBEAT_MS = 250;
/** A lock nothing refreshed inside this window is gone. */
export const LOCK_TTL_MS = 1_000;

export interface Lock {
    endpointId: string;
    sheetId: string;
    col: number;
    row: number;
    /** When the holder last said it was still there. */
    heldAt: number;
}

/**
 * One cell per peer: an editor opens on exactly one, so a new claim replaces
 * whatever that peer held before.
 */
export function claim(locks: readonly Lock[], lock: Lock): Lock[] {
    return [...locks.filter((l) => l.endpointId !== lock.endpointId), lock];
}

/** The editor closed. Instant. */
export function releaseCell(locks: readonly Lock[], endpointId: string): Lock[] {
    return locks.filter((l) => l.endpointId !== endpointId);
}

/** The connection dropped, which releases every lock that peer held. Instant. */
export function releasePeer(locks: readonly Lock[], endpointId: string): Lock[] {
    return locks.filter((l) => l.endpointId !== endpointId);
}

/** The backstop, for a frozen process on a connection that still looks alive. */
export function expire(locks: readonly Lock[], now: number, ttlMs: number = LOCK_TTL_MS): Lock[] {
    return locks.filter((l) => now - l.heldAt <= ttlMs);
}

/** Who holds this cell, if anyone. */
export function lockAt(
    locks: readonly Lock[],
    sheetId: string,
    col: number,
    row: number,
): Lock | null {
    return locks.find((l) => l.sheetId === sheetId && l.col === col && l.row === row) ?? null;
}
