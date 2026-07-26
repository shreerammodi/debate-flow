/**
 * The clock every replicated value is stamped by.
 *
 * A stamp is a hybrid logical clock: wall time, a counter that breaks ties
 * inside one millisecond, and the writing peer. Wall time is what makes "last
 * typed wins" match what the two debaters saw happen; the counter and the
 * actor make the order total, so last-writer-wins resolves identically on
 * every peer.
 */

export interface Stamp {
    /** Epoch ms, raised to the highest wall time any peer has reported. */
    ms: number;
    /** Distinguishes writes inside one millisecond. */
    counter: number;
    /** The writing peer's EndpointId. "" marks a value seeded from the file. */
    actor: string;
}

/**
 * Below every real write. Seeding uses it so two peers that open one file
 * derive byte-identical replicas without talking.
 */
export const ORIGIN_STAMP: Stamp = { ms: 0, counter: 0, actor: "" };

/** Total order: wall time, then counter, then actor. */
export function compareStamps(a: Stamp, b: Stamp): number {
    if (a.ms !== b.ms) return a.ms - b.ms;
    if (a.counter !== b.counter) return a.counter - b.counter;
    return a.actor < b.actor ? -1 : a.actor > b.actor ? 1 : 0;
}

export interface Clock {
    /** The stamp for the next local write. Strictly above every earlier one. */
    tick(): Stamp;
    /** Raise the clock past a stamp received from a peer. */
    observe(stamp: Stamp): void;
}

/**
 * `now` is injectable so tests can stall or reverse the wall clock; a stalled
 * or reversed clock must never produce a stamp that repeats or goes backwards.
 */
export function createClock(actor: string, now: () => number = Date.now): Clock {
    let ms = 0;
    let counter = 0;
    return {
        tick() {
            const wall = now();
            if (wall > ms) {
                ms = wall;
                counter = 0;
            } else {
                counter += 1;
            }
            return { ms, counter, actor };
        },
        observe(stamp) {
            if (stamp.ms > ms) {
                ms = stamp.ms;
                counter = stamp.counter;
            } else if (stamp.ms === ms && stamp.counter > counter) {
                counter = stamp.counter;
            }
        },
    };
}
