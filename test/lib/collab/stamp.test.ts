import { describe, expect, it } from "vitest";

import { compareStamps, createClock, ORIGIN_STAMP, type Stamp } from "@/lib/collab/stamp";

function fakeNow(times: number[]): () => number {
    let i = 0;
    return () => times[Math.min(i++, times.length - 1)];
}

describe("compareStamps", () => {
    it("orders by wall time first", () => {
        const early: Stamp = { ms: 1, counter: 9, actor: "z" };
        const late: Stamp = { ms: 2, counter: 0, actor: "a" };
        expect(compareStamps(early, late)).toBeLessThan(0);
        expect(compareStamps(late, early)).toBeGreaterThan(0);
    });

    it("orders by counter inside one millisecond", () => {
        const a: Stamp = { ms: 5, counter: 1, actor: "z" };
        const b: Stamp = { ms: 5, counter: 2, actor: "a" };
        expect(compareStamps(a, b)).toBeLessThan(0);
    });

    it("breaks a full tie on the actor, so every peer agrees", () => {
        const a: Stamp = { ms: 5, counter: 1, actor: "alex" };
        const b: Stamp = { ms: 5, counter: 1, actor: "sam" };
        expect(compareStamps(a, b)).toBeLessThan(0);
        expect(compareStamps(a, a)).toBe(0);
    });

    it("puts the origin stamp below any real write", () => {
        const clock = createClock("alex", () => 1_000);
        expect(compareStamps(ORIGIN_STAMP, clock.tick())).toBeLessThan(0);
    });
});

describe("createClock", () => {
    it("stamps a rising wall clock with a reset counter", () => {
        const clock = createClock("alex", fakeNow([10, 11]));
        expect(clock.tick()).toEqual({ ms: 10, counter: 0, actor: "alex" });
        expect(clock.tick()).toEqual({ ms: 11, counter: 0, actor: "alex" });
    });

    it("stays strictly increasing when the wall clock stalls", () => {
        const clock = createClock("alex", () => 10);
        expect(clock.tick()).toEqual({ ms: 10, counter: 0, actor: "alex" });
        expect(clock.tick()).toEqual({ ms: 10, counter: 1, actor: "alex" });
        expect(clock.tick()).toEqual({ ms: 10, counter: 2, actor: "alex" });
    });

    it("stays strictly increasing when the wall clock walks backwards", () => {
        const clock = createClock("alex", fakeNow([20, 5]));
        const first = clock.tick();
        const second = clock.tick();
        expect(compareStamps(first, second)).toBeLessThan(0);
        expect(second).toEqual({ ms: 20, counter: 1, actor: "alex" });
    });

    it("raises past a remote stamp so the next local write beats it", () => {
        const clock = createClock("alex", () => 10);
        clock.observe({ ms: 900, counter: 4, actor: "sam" });
        const next = clock.tick();
        expect(compareStamps({ ms: 900, counter: 4, actor: "sam" }, next)).toBeLessThan(0);
        expect(next).toEqual({ ms: 900, counter: 5, actor: "alex" });
    });

    it("ignores a remote stamp that is already behind", () => {
        const clock = createClock("alex", () => 100);
        clock.tick();
        clock.observe({ ms: 1, counter: 0, actor: "sam" });
        expect(clock.tick()).toEqual({ ms: 100, counter: 1, actor: "alex" });
    });

    it("raises past a count no clock reported instead of ignoring the stamp", () => {
        const clock = createClock("alex", () => 10);
        // A peer sitting at the far end of the safe range with a count a
        // million times past what a stalled clock reaches. Dropping the whole
        // stamp would leave this clock at 10, and every later local write
        // losing to it forever.
        const pinned = { ms: 9_007_199_254_740_991, counter: 2_000_000, actor: "sam" };
        clock.observe(pinned);
        expect(compareStamps(pinned, clock.tick())).toBeLessThan(0);
    });
});
