import { describe, expect, it } from "vitest";

import { rankBetween, seedRank } from "@/lib/collab/rank";

describe("seedRank", () => {
    it("is fixed width, so a plain string sort is a row sort", () => {
        const ranks = [0, 1, 2, 30, 61, 500].map(seedRank);
        expect(new Set(ranks.map((r) => r.length)).size).toBe(1);
        expect(ranks.slice().sort()).toEqual(ranks);
    });

    it("derives the same rank for a row on every peer", () => {
        expect(seedRank(7)).toBe(seedRank(7));
    });

    it("never ends in the zero digit", () => {
        for (let i = 0; i < 200; i++) expect(seedRank(i).endsWith("0")).toBe(false);
    });

    it("leaves room between neighbouring rows", () => {
        expect(rankBetween(seedRank(0), seedRank(1))).not.toBe(seedRank(0));
    });
});

describe("rankBetween", () => {
    it("returns a rank strictly between two seeds", () => {
        const a = seedRank(3);
        const b = seedRank(4);
        const mid = rankBetween(a, b);
        expect(a < mid).toBe(true);
        expect(mid < b).toBe(true);
    });

    it("returns a rank above the last row when there is no successor", () => {
        const last = seedRank(9);
        expect(rankBetween(last, null) > last).toBe(true);
    });

    it("returns a rank below the first row when there is no predecessor", () => {
        const first = seedRank(0);
        expect(rankBetween(null, first) < first).toBe(true);
    });

    it("returns a rank for an empty column", () => {
        const only = rankBetween(null, null);
        expect(only.length).toBeGreaterThan(0);
        expect(rankBetween(null, only) < only).toBe(true);
    });

    it("subdivides one gap two hundred times and stays ordered", () => {
        let low = seedRank(0);
        const high = seedRank(1);
        const seen: string[] = [];
        for (let i = 0; i < 200; i++) {
            const next = rankBetween(low, high);
            expect(low < next).toBe(true);
            expect(next < high).toBe(true);
            seen.push(next);
            low = next;
        }
        expect(new Set(seen).size).toBe(200);
        expect(seen.slice().sort()).toEqual(seen);
    });

    it("appends two hundred times at the end and stays ordered", () => {
        let last = seedRank(0);
        const seen: string[] = [];
        for (let i = 0; i < 200; i++) {
            const next = rankBetween(last, null);
            expect(last < next).toBe(true);
            seen.push(next);
            last = next;
        }
        expect(seen.slice().sort()).toEqual(seen);
    });

    it("refuses a pair that is out of order", () => {
        expect(() => rankBetween(seedRank(4), seedRank(3))).toThrow();
    });
});
