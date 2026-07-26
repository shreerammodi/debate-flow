import { describe, expect, it } from "vitest";

import { ORIGIN_STAMP } from "@/lib/collab/stamp";
import {
    cellKey,
    compareCells,
    flattenLeaves,
    setPath,
    type CollabCell,
    type Json,
} from "@/lib/collab/types";

function cell(col: number, rank: string, actor: string): CollabCell {
    return {
        col,
        rank,
        actor,
        text: null,
        textStamp: ORIGIN_STAMP,
        meta: {},
        metaStamp: ORIGIN_STAMP,
        deleted: null,
    };
}

describe("cellKey", () => {
    it("identifies a cell by column, rank, and creator", () => {
        expect(cellKey(2, "V", "alex")).toBe(cellKey(2, "V", "alex"));
        expect(cellKey(2, "V", "alex")).not.toBe(cellKey(2, "V", "sam"));
        expect(cellKey(2, "V", "alex")).not.toBe(cellKey(3, "V", "alex"));
    });
});

describe("compareCells", () => {
    it("orders by rank, not by the concatenated key", () => {
        // "A" sorts below "A0", which a key that joins on a separator would
        // get backwards.
        expect(compareCells(cell(0, "A", "z"), cell(0, "A0", "a"))).toBeLessThan(0);
    });

    it("breaks a rank tie on the creator, so both cells survive in one order", () => {
        expect(compareCells(cell(0, "V", "alex"), cell(0, "V", "sam"))).toBeLessThan(0);
        expect(compareCells(cell(0, "V", "sam"), cell(0, "V", "alex"))).toBeGreaterThan(0);
        expect(compareCells(cell(0, "V", "alex"), cell(0, "V", "alex"))).toBe(0);
    });
});

describe("leaf paths", () => {
    it("flattens a nested object into dotted leaves", () => {
        const out: Record<string, Json> = {};
        flattenLeaves({ event: "pf", scouting: { aff: { first: { last: "Ito" } } } }, "", out);
        expect(out).toEqual({ event: "pf", "scouting.aff.first.last": "Ito" });
    });

    it("treats an array as one leaf", () => {
        const out: Record<string, Json> = {};
        flattenLeaves({ tags: ["a", "b"] }, "", out);
        expect(out).toEqual({ tags: ["a", "b"] });
    });

    it("skips an undefined leaf so an absent field stays absent", () => {
        const out: Record<string, Json> = {};
        flattenLeaves({ a: 1, b: undefined }, "", out);
        expect(out).toEqual({ a: 1 });
    });

    it("rebuilds the object the leaves came from", () => {
        const source = {
            event: "pf",
            scouting: { tournament: "Harvard", decision: { vote: "aff" } },
        };
        const leaves: Record<string, Json> = {};
        flattenLeaves(source, "", leaves);
        const rebuilt: Record<string, unknown> = {};
        for (const [path, value] of Object.entries(leaves)) setPath(rebuilt, path, value);
        expect(rebuilt).toEqual(source);
    });

    it("keeps a path a newer build wrote and this one does not read", () => {
        const rebuilt: Record<string, unknown> = {};
        setPath(rebuilt, "scouting.decision.peerNotes.sam", "vote neg");
        expect(rebuilt).toEqual({ scouting: { decision: { peerNotes: { sam: "vote neg" } } } });
    });
});
