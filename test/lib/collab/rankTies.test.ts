import { describe, expect, it } from "vitest";

import { liveCells, seedDoc } from "@/lib/collab/doc";
import { merge } from "@/lib/collab/merge";
import { applyOp, type OpContext } from "@/lib/collab/ops";
import { createClock } from "@/lib/collab/stamp";
import type { CollabDoc } from "@/lib/collab/types";
import { makeFlowRound, type FlowRound } from "@/lib/model/flow";

/**
 * Two peers inserting at one position derive the same rank, and the merge keeps
 * both cells with their authors breaking the tie. That is the design. What
 * follows is what happens to the next insert aimed at the pair, where there is
 * no rank left to sit between them.
 */

function ctxFor(actor: string, startMs: number): OpContext {
    let t = startMs;
    return { actor, clock: createClock(actor, () => t++) };
}

function round(): FlowRound {
    const r = makeFlowRound({});
    for (const sheet of r.sheets) {
        sheet.data = [["perm"], ["cap bad"], ["extend"]];
    }
    return r;
}

/** A column where two peers inserted at row 1 at the same moment. */
function tied(): { doc: CollabDoc; sheetId: string; base: FlowRound } {
    const base = round();
    const sheetId = base.sheets.find((s) => s.kind !== "cx")!.id;
    const seeded = seedDoc(base);
    const alex = applyOp(
        seeded,
        { kind: "insertCell", sheetId, col: 0, row: 1 },
        ctxFor("alex", 1_000),
    );
    const sam = applyOp(
        seeded,
        { kind: "insertCell", sheetId, col: 0, row: 1 },
        ctxFor("sam", 5_000),
    );
    return { doc: merge(alex, sam).doc, sheetId, base };
}

describe("two cells that share a rank", () => {
    it("both survive the merge, in a deterministic order", () => {
        const { doc, sheetId } = tied();
        const column = liveCells(doc.sheets[sheetId], 0);
        expect(column).toHaveLength(5);
        expect(column[1].rank).toBe(column[2].rank);
        expect(column[1].actor < column[2].actor).toBe(true);
    });

    it("takes another insert aimed between them without throwing", () => {
        const { doc, sheetId } = tied();
        const kim = ctxFor("kim", 9_000);
        expect(() =>
            applyOp(doc, { kind: "insertCell", sheetId, col: 0, row: 2 }, kim),
        ).not.toThrow();
    });

    it("takes an insert at every row of the tied column", () => {
        const { doc, sheetId } = tied();
        const kim = ctxFor("kim", 9_000);
        const height = liveCells(doc.sheets[sheetId], 0).length;
        for (let row = 0; row <= height; row++) {
            expect(
                () => applyOp(doc, { kind: "insertCell", sheetId, col: 0, row }, kim),
                `insert at row ${row}`,
            ).not.toThrow();
        }
    });

    it("takes a whole-row insert across a sheet holding a tie", () => {
        const { doc, sheetId } = tied();
        const kim = ctxFor("kim", 9_000);
        expect(() => applyOp(doc, { kind: "insertRow", sheetId, row: 2 }, kim)).not.toThrow();
    });

    it("puts the new cell somewhere, and keeps every cell that was there", () => {
        const { doc, sheetId } = tied();
        const kim = ctxFor("kim", 9_000);
        const before = liveCells(doc.sheets[sheetId], 0);
        const after = applyOp(doc, { kind: "insertCell", sheetId, col: 0, row: 2 }, kim);
        const column = liveCells(after.sheets[sheetId], 0);

        expect(column).toHaveLength(before.length + 1);
        // Nothing that existed moved out from under anyone or changed identity.
        const keyOf = (c: { rank: string; actor: string }) => `${c.rank}|${c.actor}`;
        for (const cell of before) {
            expect(column.map(keyOf)).toContain(keyOf(cell));
        }
    });

    it("stays a tie the two peers agree on, whichever applied the insert", () => {
        const { doc, sheetId } = tied();
        const kim = ctxFor("kim", 9_000);
        const mine = applyOp(doc, { kind: "insertCell", sheetId, col: 0, row: 2 }, kim);
        // The far side receives it and merges rather than deriving it.
        const theirs = merge(doc, mine).doc;
        expect(liveCells(theirs.sheets[sheetId], 0).map((c) => `${c.rank}|${c.actor}`)).toEqual(
            liveCells(mine.sheets[sheetId], 0).map((c) => `${c.rank}|${c.actor}`),
        );
    });
});
