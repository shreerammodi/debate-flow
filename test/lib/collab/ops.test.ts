import { describe, expect, it } from "vitest";

import { liveCells, projectDoc, seedDoc } from "@/lib/collab/doc";
import { applyOp, type CollabOp, type OpContext } from "@/lib/collab/ops";
import { createClock } from "@/lib/collab/stamp";
import type { CollabDoc } from "@/lib/collab/types";
import { makeFlowRound, makeFlowSheet, type FlowRound } from "@/lib/model/flow";

function fixture(): { round: FlowRound; doc: CollabDoc; sheetId: string; ctx: OpContext } {
    const round = makeFlowRound({});
    const flow = round.sheets.find((s) => s.kind !== "cx")!;
    flow.data = [
        ["perm", "link"],
        ["cap bad", "turn"],
    ];
    flow.meta = {};
    let t = 1_000;
    return {
        round,
        doc: seedDoc(round),
        sheetId: flow.id,
        ctx: { actor: "alex", clock: createClock("alex", () => t++) },
    };
}

function grid(doc: CollabDoc, round: FlowRound, sheetId: string): (string | null)[][] {
    return projectDoc(doc, round).sheets.find((s) => s.id === sheetId)!.data;
}

function run(doc: CollabDoc, ctx: OpContext, ops: CollabOp[]): CollabDoc {
    return ops.reduce((acc, op) => applyOp(acc, op, ctx), doc);
}

describe("applyOp", () => {
    it("never mutates the document it is given", () => {
        const { doc, sheetId, ctx } = fixture();
        const before = structuredClone(doc);
        applyOp(doc, { kind: "cellText", sheetId, col: 0, row: 0, text: "changed" }, ctx);
        expect(doc).toEqual(before);
    });

    it("writes a cell at a grid coordinate", () => {
        const { doc, round, sheetId, ctx } = fixture();
        const next = applyOp(doc, { kind: "cellText", sheetId, col: 1, row: 1, text: "x" }, ctx);
        expect(grid(next, round, sheetId)).toEqual([
            ["perm", "link"],
            ["cap bad", "x"],
        ]);
    });

    it("grows a column when the grid writes past its last row", () => {
        const { doc, round, sheetId, ctx } = fixture();
        const next = applyOp(doc, { kind: "cellText", sheetId, col: 0, row: 3, text: "deep" }, ctx);
        expect(grid(next, round, sheetId)).toEqual([
            ["perm", "link"],
            ["cap bad", "turn"],
            [null, null],
            ["deep", null],
        ]);
    });

    it("writes meta without disturbing the text stamp", () => {
        const { doc, round, sheetId, ctx } = fixture();
        const textStampBefore = liveCells(doc.sheets[sheetId], 0)[0].textStamp;
        const next = applyOp(
            doc,
            { kind: "cellMeta", sheetId, col: 0, row: 0, meta: { bold: true } },
            ctx,
        );
        const cell = liveCells(next.sheets[sheetId], 0)[0];
        expect(cell.meta).toEqual({ bold: true });
        expect(cell.text).toBe("perm");
        expect(cell.textStamp).toEqual(textStampBefore);
        expect(projectDoc(next, round).sheets.find((s) => s.id === sheetId)!.meta).toEqual({
            "0,0": { bold: true },
        });
    });

    it("inserts into one column and leaves the others where they were", () => {
        const { doc, round, sheetId, ctx } = fixture();
        const next = applyOp(doc, { kind: "insertCell", sheetId, col: 0, row: 1 }, ctx);
        expect(grid(next, round, sheetId)).toEqual([
            ["perm", "link"],
            [null, "turn"],
            ["cap bad", null],
        ]);
    });

    it("removes one cell and closes the gap in its column only", () => {
        const { doc, round, sheetId, ctx } = fixture();
        const next = applyOp(doc, { kind: "removeCell", sheetId, col: 1, row: 0 }, ctx);
        expect(grid(next, round, sheetId)).toEqual([
            ["perm", "turn"],
            ["cap bad", null],
        ]);
    });

    it("inserts a row across every column", () => {
        const { doc, round, sheetId, ctx } = fixture();
        const next = applyOp(doc, { kind: "insertRow", sheetId, row: 1 }, ctx);
        expect(grid(next, round, sheetId)).toEqual([
            ["perm", "link"],
            [null, null],
            ["cap bad", "turn"],
        ]);
    });

    it("removes a row across every column", () => {
        const { doc, round, sheetId, ctx } = fixture();
        const next = applyOp(doc, { kind: "removeRow", sheetId, row: 0 }, ctx);
        expect(grid(next, round, sheetId)).toEqual([["cap bad", "turn"]]);
    });

    it("keeps a tombstone so a peer that still holds the cell cannot revive it", () => {
        const { doc, sheetId, ctx } = fixture();
        const next = applyOp(doc, { kind: "removeRow", sheetId, row: 0 }, ctx);
        const tombstones = Object.values(next.sheets[sheetId].cells).filter((c) => c.deleted);
        expect(tombstones).toHaveLength(2);
    });

    it("writes a sheet field and a round field as registers", () => {
        const { doc, round, sheetId, ctx } = fixture();
        const next = run(doc, ctx, [
            { kind: "sheetField", sheetId, path: "title", value: "T" },
            { kind: "roundField", path: "scouting.tournament", value: "Harvard" },
        ]);
        const projected = projectDoc(next, round);
        expect(projected.sheets.find((s) => s.id === sheetId)!.title).toBe("T");
        expect(projected.scouting.tournament).toBe("Harvard");
    });

    it("adds and removes a sheet", () => {
        const { doc, round, ctx } = fixture();
        const added = makeFlowSheet({ title: "DA", group: "neg", order: 5 });
        const next = applyOp(doc, { kind: "addSheet", sheet: added }, ctx);
        expect(projectDoc(next, round).sheets.map((s) => s.title)).toContain("DA");
        const gone = applyOp(next, { kind: "removeSheet", sheetId: added.id }, ctx);
        expect(projectDoc(gone, round).sheets.map((s) => s.title)).not.toContain("DA");
    });

    it("ignores an op aimed at a sheet that is not here", () => {
        const { doc, ctx } = fixture();
        expect(
            applyOp(doc, { kind: "cellText", sheetId: "gone", col: 0, row: 0, text: "x" }, ctx),
        ).toEqual(doc);
    });

    it("keeps two concurrent inserts at one row, in a stable order", () => {
        const { doc, round, sheetId } = fixture();
        let ta = 2_000;
        let tb = 2_000;
        const alex = { actor: "alex", clock: createClock("alex", () => ta++) };
        const sam = { actor: "sam", clock: createClock("sam", () => tb++) };
        const fromAlex = applyOp(doc, { kind: "insertCell", sheetId, col: 0, row: 1 }, alex);
        const fromSam = applyOp(doc, { kind: "insertCell", sheetId, col: 0, row: 1 }, sam);
        const both = {
            ...fromAlex,
            sheets: {
                ...fromAlex.sheets,
                [sheetId]: {
                    ...fromAlex.sheets[sheetId],
                    cells: { ...fromAlex.sheets[sheetId].cells, ...fromSam.sheets[sheetId].cells },
                },
            },
        };
        expect(grid(both, round, sheetId)).toEqual([
            ["perm", "link"],
            [null, "turn"],
            [null, null],
            ["cap bad", null],
        ]);
        expect(liveCells(both.sheets[sheetId], 0).map((c) => c.actor)).toEqual([
            "",
            "alex",
            "sam",
            "",
        ]);
    });
});
