import { describe, expect, it, vi } from "vitest";

import { BACKOFF_CEILING_MS, backoffMs, retryForever } from "@/lib/collab/reconnect";

function manualClock() {
    let pending: { fn: () => void; at: number }[] = [];
    let now = 0;
    return {
        schedule(fn: () => void, ms: number) {
            const entry = { fn, at: now + ms };
            pending.push(entry);
            return () => {
                pending = pending.filter((p) => p !== entry);
            };
        },
        delays(): number[] {
            return pending.map((p) => p.at - now);
        },
        async advance(ms: number) {
            now += ms;
            const due = pending.filter((p) => p.at <= now);
            pending = pending.filter((p) => p.at > now);
            for (const p of due) p.fn();
            // Let the dial promise settle before the test looks.
            await Promise.resolve();
            await Promise.resolve();
        },
        get pendingCount() {
            return pending.length;
        },
    };
}

describe("backoffMs", () => {
    it("starts short, so a blink costs nothing", () => {
        expect(backoffMs(0, () => 0.5)).toBeLessThanOrEqual(1_000);
    });

    it("doubles as attempts fail", () => {
        const mid = () => 0.5;
        expect(backoffMs(2, mid)).toBeGreaterThan(backoffMs(1, mid));
        expect(backoffMs(3, mid)).toBeGreaterThan(backoffMs(2, mid));
    });

    it("clamps, so a long outage still retries on a human cadence", () => {
        for (const attempt of [10, 20, 100]) {
            expect(backoffMs(attempt, () => 1)).toBeLessThanOrEqual(BACKOFF_CEILING_MS);
        }
    });

    it("jitters, so two peers dropped together do not retry in lockstep", () => {
        expect(backoffMs(5, () => 0)).not.toBe(backoffMs(5, () => 1));
    });

    it("never returns a negative or zero delay", () => {
        for (let attempt = 0; attempt < 12; attempt++) {
            expect(backoffMs(attempt, () => 0)).toBeGreaterThan(0);
        }
    });
});

describe("retryForever", () => {
    it("keeps dialling a peer that stays down", async () => {
        const clock = manualClock();
        const dial = vi.fn().mockRejectedValue(new Error("down"));
        retryForever({ dial, schedule: clock.schedule, random: () => 0.5 });

        await clock.advance(60_000);
        await clock.advance(60_000);
        await clock.advance(60_000);
        expect(dial.mock.calls.length).toBeGreaterThanOrEqual(3);
    });

    it("backs off further each time it fails", async () => {
        const clock = manualClock();
        const dial = vi.fn().mockRejectedValue(new Error("down"));
        retryForever({ dial, schedule: clock.schedule, random: () => 0.5 });

        const first = clock.delays()[0];
        await clock.advance(first);
        const second = clock.delays()[0];
        await clock.advance(second);
        expect(clock.delays()[0]).toBeGreaterThan(first);
    });

    it("resets after a success, so a second drop does not start at the ceiling", async () => {
        const clock = manualClock();
        const dial = vi
            .fn()
            .mockRejectedValueOnce(new Error("down"))
            .mockRejectedValueOnce(new Error("down"))
            .mockResolvedValueOnce(undefined);
        const retry = retryForever({ dial, schedule: clock.schedule, random: () => 0.5 });

        await clock.advance(clock.delays()[0]);
        await clock.advance(clock.delays()[0]);
        await clock.advance(clock.delays()[0]);
        // A success stops the schedule; the caller re-arms on the next drop.
        expect(clock.pendingCount).toBe(0);

        retry.stop();
    });

    it("stops dead when told to", async () => {
        const clock = manualClock();
        const dial = vi.fn().mockRejectedValue(new Error("down"));
        const retry = retryForever({ dial, schedule: clock.schedule, random: () => 0.5 });

        retry.stop();
        await clock.advance(120_000);
        expect(dial).not.toHaveBeenCalled();
        expect(clock.pendingCount).toBe(0);
    });

    it("never rejects, because a dead peer is ordinary", async () => {
        const clock = manualClock();
        const dial = vi.fn().mockRejectedValue(new Error("down"));
        expect(() =>
            retryForever({ dial, schedule: clock.schedule, random: () => 0.5 }),
        ).not.toThrow();
        await clock.advance(60_000);
    });
});
