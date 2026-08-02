import { beforeEach, describe, expect, it } from "vitest";

import { liveCells, seedDoc } from "@/lib/collab/doc";
import { gridPatchFor } from "@/lib/collab/gridPatch";
import { applyOp, type CollabOp, type OpContext } from "@/lib/collab/ops";
import { createClock } from "@/lib/collab/stamp";
import type { CollabDoc, CollabSheet } from "@/lib/collab/types";
import { modelCol } from "@/lib/grid/colSpace";
import { makeFlowRound, type FlowRound } from "@/lib/model/flow";

let base: FlowRound;
let sheetId: string;
let before: CollabDoc;
let sam: OpContext;

beforeEach(() => {
    base = makeFlowRound({});
    const sheet = base.sheets.find((s) => s.kind !== "cx")!;
    sheet.data = [
        ["perm", "link"],
        ["cap bad", "turn"],
    ];
    sheetId = sheet.id;
    before = seedDoc(base);
    let t = 5_000;
    sam = { actor: "sam", clock: createClock("sam", () => t++) };
});

function after(...ops: CollabOp[]): CollabSheet {
    let doc = before;
    for (const op of ops) doc = applyOp(doc, op, sam);
    return doc.sheets[sheetId];
}

describe("the grid writes a partner's change comes down to", () => {
    it("names only the cell whose text moved", () => {
        const patch = gridPatchFor(
            before.sheets[sheetId],
            after({ kind: "cellText", sheetId, col: 1, row: 0, text: "no link" }),
        );
        expect(patch.writes).toEqual([{ row: 0, col: 1, text: "no link" }]);
        expect(patch.meta).toEqual([]);
        expect(patch.height).toBe(2);
    });

    it("names nothing at all when a delta changed no cell", () => {
        const patch = gridPatchFor(before.sheets[sheetId], before.sheets[sheetId]);
        expect(patch.writes).toEqual([]);
        expect(patch.meta).toEqual([]);
    });

    it("carries a decoration without touching the text", () => {
        const patch = gridPatchFor(
            before.sheets[sheetId],
            after({ kind: "cellMeta", sheetId, col: 0, row: 1, meta: { bold: true } }),
        );
        expect(patch.writes).toEqual([]);
        expect(patch.meta).toEqual([{ row: 1, col: 0, meta: { bold: true } }]);
    });

    it("clears a decoration a partner removed", () => {
        const bolded = after({ kind: "cellMeta", sheetId, col: 0, row: 0, meta: { bold: true } });
        const plain = applyOp(
            { ...before, sheets: { ...before.sheets, [sheetId]: bolded } },
            { kind: "cellMeta", sheetId, col: 0, row: 0, meta: {} },
            sam,
        ).sheets[sheetId];
        expect(gridPatchFor(bolded, plain).meta).toEqual([{ row: 0, col: 0, meta: null }]);
    });

    it("rewrites the rows a partner's insert pushed down, and grows the height", () => {
        const patch = gridPatchFor(
            before.sheets[sheetId],
            after(
                { kind: "insertCell", sheetId, col: 0, row: 0 },
                { kind: "cellText", sheetId, col: 0, row: 0, text: "extend" },
            ),
        );
        expect(patch.writes).toEqual([
            { row: 0, col: 0, text: "extend" },
            { row: 1, col: 0, text: "perm" },
            { row: 2, col: 0, text: "cap bad" },
        ]);
        expect(patch.height).toBe(3);
    });

    it("blanks the tail a partner's row removal left behind", () => {
        const patch = gridPatchFor(
            before.sheets[sheetId],
            after({ kind: "removeRow", sheetId, row: 0 }),
        );
        expect(patch.writes).toEqual([
            { row: 0, col: 0, text: "cap bad" },
            { row: 1, col: 0, text: null },
            { row: 0, col: 1, text: "turn" },
            { row: 1, col: 1, text: null },
        ]);
    });

    it("leaves out the cell the editor is open on, and no other", () => {
        const next = after(
            { kind: "cellText", sheetId, col: 0, row: 0, text: "theirs" },
            { kind: "cellText", sheetId, col: 0, row: 1, text: "also theirs" },
        );
        const held = liveCells(next, 0)[0];
        const patch = gridPatchFor(before.sheets[sheetId], next, [
            { col: modelCol(0), rank: held.rank, actor: held.actor },
        ]);
        expect(patch.writes).toEqual([{ row: 1, col: 0, text: "also theirs" }]);
    });

    it("writes every cell of a sheet a partner just added", () => {
        const patch = gridPatchFor(undefined, before.sheets[sheetId]);
        expect(patch.writes).toHaveLength(4);
        expect(patch.height).toBe(2);
    });
});
