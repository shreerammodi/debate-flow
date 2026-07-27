import { beforeEach, describe, expect, it } from "vitest";

import { seedDoc } from "@/lib/collab/doc";
import { applyOp, type OpContext } from "@/lib/collab/ops";
import { followSelection, rowOfIdentity, selectionIdentity } from "@/lib/collab/selection";
import { createClock } from "@/lib/collab/stamp";
import type { CollabDoc, CollabSheet } from "@/lib/collab/types";
import { makeFlowRound, type FlowRound } from "@/lib/model/flow";

let round: FlowRound;
let sheetId: string;
let doc: CollabDoc;
let sam: OpContext;

beforeEach(() => {
    round = makeFlowRound({});
    const flow = round.sheets.find((s) => s.kind !== "cx")!;
    sheetId = flow.id;
    flow.data = [
        ["a0", "b0"],
        ["a1", "b1"],
        ["a2", "b2"],
        ["a3", "b3"],
    ];
    doc = seedDoc(round);
    let t = 9_000;
    sam = { actor: "sam", clock: createClock("sam", () => t++) };
});

function sheet(d: CollabDoc): CollabSheet {
    return d.sheets[sheetId];
}

/** What a partner did, as the pair of sheets the selection has to survive. */
function remote(...ops: Parameters<typeof applyOp>[1][]): [CollabSheet, CollabSheet] {
    const before = sheet(doc);
    let next = doc;
    for (const op of ops) next = applyOp(next, op, sam);
    return [before, sheet(next)];
}

describe("selectionIdentity", () => {
    it("names the cell under the cursor by something no insert can change", () => {
        const ref = selectionIdentity(sheet(doc), 2, 0)!;
        expect(ref.col).toBe(0);
        expect(ref.rank).toBeTruthy();
        expect(rowOfIdentity(sheet(doc), ref)).toBe(2);
    });

    it("has nothing to name past the end of a column", () => {
        expect(selectionIdentity(sheet(doc), 99, 0)).toBeNull();
    });

    it("has nothing to name in a column that holds no cells", () => {
        expect(selectionIdentity(sheet(doc), 0, 7)).toBeNull();
    });
});

describe("followSelection", () => {
    it("moves down one when a partner inserts a row above the cursor", () => {
        const [before, after] = remote({ kind: "insertRow", sheetId, row: 1 });
        expect(followSelection(before, after, 2, 0)).toBe(3);
    });

    it("moves down two when a partner inserts two rows above", () => {
        const [before, after] = remote(
            { kind: "insertRow", sheetId, row: 0 },
            { kind: "insertRow", sheetId, row: 0 },
        );
        expect(followSelection(before, after, 2, 0)).toBe(4);
    });

    it("stays put when a partner inserts below the cursor", () => {
        const [before, after] = remote({ kind: "insertRow", sheetId, row: 3 });
        expect(followSelection(before, after, 1, 0)).toBe(1);
    });

    it("moves up one when a partner deletes a row above the cursor", () => {
        const [before, after] = remote({ kind: "removeRow", sheetId, row: 0 });
        expect(followSelection(before, after, 2, 0)).toBe(1);
    });

    it("holds its index when a partner deletes the row the cursor sits in", () => {
        const [before, after] = remote({ kind: "removeRow", sheetId, row: 2 });
        expect(followSelection(before, after, 2, 0)).toBe(2);
    });

    it("moves down one when a partner inserts a cell above in the same column", () => {
        const [before, after] = remote({ kind: "insertCell", sheetId, col: 0, row: 1 });
        expect(followSelection(before, after, 2, 0)).toBe(3);
    });

    it("stays put when the insert lands in another column", () => {
        const [before, after] = remote({ kind: "insertCell", sheetId, col: 1, row: 0 });
        expect(followSelection(before, after, 2, 0)).toBe(2);
    });

    it("stays put for a write that moves nothing", () => {
        const [before, after] = remote({
            kind: "cellText",
            sheetId,
            col: 0,
            row: 0,
            text: "typed",
        });
        expect(followSelection(before, after, 2, 0)).toBe(2);
    });

    it("holds an index it cannot name, rather than guessing", () => {
        const [before, after] = remote({ kind: "insertRow", sheetId, row: 0 });
        expect(followSelection(before, after, 99, 0)).toBe(99);
        expect(followSelection(before, after, 0, 7)).toBe(0);
    });
});
