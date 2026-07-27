/**
 * Where the advisory locks a session holds reach the live grid.
 *
 * Locks are grid state, so they live beside the other grid registries rather
 * than in `useCollabStore`. Each peer refreshes its lock on a 250ms heartbeat,
 * and the store is what the session chip subscribes to; routing presence
 * through it would re-render React several times a second for a value only the
 * grid paints.
 *
 * A lock landing here has to repaint something, so the registry notifies as
 * well as stores. Expiry is deliberately not its business: `lockClassFor`
 * takes a `now` and asks presence, so a stale lock stops painting on the next
 * render with no timer here.
 */

import type { Lock } from "@/lib/collab/presence";

/** One array backs every empty lock table, so a clear changes no identity. */
const NO_LOCKS: readonly Lock[] = [];

let locks: readonly Lock[] = NO_LOCKS;
const listeners = new Set<() => void>();

/** The session layer publishes the whole table; there is no incremental path. */
export function setLocks(next: readonly Lock[]): void {
    locks = next.length === 0 ? NO_LOCKS : next;
    for (const listener of listeners) listener();
}

export function getLocks(): readonly Lock[] {
    return locks;
}

/**
 * A mounted grid asks to repaint when the table changes. Both panes of a split
 * subscribe, so listeners are a set and the returned unsubscribe drops only
 * the caller's own.
 */
export function onLocksChanged(cb: () => void): () => void {
    listeners.add(cb);
    return () => {
        listeners.delete(cb);
    };
}
