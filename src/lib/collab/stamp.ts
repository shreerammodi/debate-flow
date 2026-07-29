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

/**
 * Writes inside one millisecond. A counter near the float limit cannot be
 * climbed past: `tick` moves it by one, so once the clock has adopted such a
 * counter every later local write carries the same stamp, `known` reads each of
 * them as one the far side already holds, and the debater's edits stop shipping.
 */
const MAX_COUNTER = 1_000_000;

/**
 * Total order: wall time, then counter, then actor.
 *
 * Both counts come off the wire inside a peer's document and nothing above this
 * narrows them. Subtraction on `NaN` yields `NaN`, every comparison against
 * zero is false, and each caller reads that as "the other side is greater", so
 * one unusable count would win every comparison here forever. A count that is
 * not one sorts at the origin instead, below every real write.
 */
export function compareStamps(a: Stamp, b: Stamp): number {
    const ams = Number.isSafeInteger(a.ms) ? a.ms : 0;
    const bms = Number.isSafeInteger(b.ms) ? b.ms : 0;
    if (ams !== bms) return ams - bms;
    const ac = Number.isSafeInteger(a.counter) ? a.counter : 0;
    const bc = Number.isSafeInteger(b.counter) ? b.counter : 0;
    if (ac !== bc) return ac - bc;
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
            // A peer's clock is the peer's to set and this one adopts its wall
            // time, which is what makes "last typed wins" match what the two
            // debaters saw. A reading no clock produces is not adopted at all.
            if (!Number.isSafeInteger(stamp.ms) || stamp.ms < 0) return;
            if (!Number.isSafeInteger(stamp.counter) || stamp.counter < 0) return;
            if (stamp.counter > MAX_COUNTER) return;
            if (stamp.ms > ms) {
                ms = stamp.ms;
                counter = stamp.counter;
            } else if (stamp.ms === ms && stamp.counter > counter) {
                counter = stamp.counter;
            }
        },
    };
}
