import { beforeEach, describe, expect, it } from "vitest";

import { liveCells, projectSheet, seedDoc } from "@/lib/collab/doc";
import { rowOpFromHook } from "@/lib/collab/gridOps";
import { merge } from "@/lib/collab/merge";
import { applyOp, type CollabOp, type OpContext } from "@/lib/collab/ops";
import { openSpanOps } from "@/lib/collab/spanOps";
import { createClock } from "@/lib/collab/stamp";
import type { CollabDoc } from "@/lib/collab/types";
import { makeFlowRound, type FlowRound } from "@/lib/model/flow";

/**
 * An insert-mode paste and a CardMirror send both open a run in one column.
 *
 * Both used to re-derive the whole sheet afterwards, which re-keys every cell
 * from its row position: invisible on one machine, and on two the peer holds
 * keys that name nothing, merges the two sets, and the column doubles. Said as
 * ops, the partner applies the same thing.
 */

let round: FlowRound;
let sheetId: string;
let base: CollabDoc;

function ctxFor(actor: string, startMs: number): OpContext {
    let t = startMs;
    return { actor, clock: createClock(actor, () => t++) };
}

function run(doc: CollabDoc, ctx: OpContext, ops: CollabOp[]): CollabDoc {
    let next = doc;
    for (const op of ops) next = applyOp(next, op, ctx);
    return next;
}

function column(doc: CollabDoc, col: number): (string | null)[] {
    return projectSheet(doc.sheets[sheetId]).data.map((r) => r[col] ?? null);
}

beforeEach(() => {
    round = makeFlowRound({});
    const sheet = round.sheets.find((s) => s.kind !== "cx")!;
    sheet.data = [
        ["perm", "link"],
        ["cap bad", "turn"],
        ["extend", "drop"],
    ];
    sheetId = sheet.id;
    base = seedDoc(round);
});

describe("openSpanOps", () => {
    it("opens the run from the top, pushing the tail down", () => {
        const alex = ctxFor("alex", 1_000);
        const opened = run(base, alex, openSpanOps(sheetId, 0, 1, 2));
        expect(column(opened, 0)).toEqual(["perm", null, null, "cap bad", "extend"]);
        // The neighbouring column is untouched: a paste lands where it lands.
        expect(column(opened, 1)).toEqual(["link", "turn", "drop", null, null]);
    });

    it("carries each cell's decoration down with it", () => {
        const alex = ctxFor("alex", 1_000);
        const bolded = run(base, alex, [
            { kind: "cellMeta", sheetId, col: 0, row: 1, meta: { bold: true } },
        ]);
        const opened = run(bolded, alex, openSpanOps(sheetId, 0, 1, 1));
        expect(projectSheet(opened.sheets[sheetId]).meta["2,0"]).toEqual({ bold: true });
        expect(projectSheet(opened.sheets[sheetId]).meta["1,0"]).toBeUndefined();
    });

    it("opens nothing for a run of no rows", () => {
        expect(openSpanOps(sheetId, 0, 0, 0)).toEqual([]);
    });
});

describe("a paste both machines see", () => {
    it("leaves the two replicas holding the same column", () => {
        const alex = ctxFor("alex", 1_000);
        const sam = ctxFor("sam", 5_000);

        // Alex pastes two rows in at row 1, the way beforePaste records it,
        // then the text lands the way afterChange records it.
        let mine = run(base, alex, openSpanOps(sheetId, 0, 1, 2));
        mine = run(mine, alex, [
            { kind: "cellText", sheetId, col: 0, row: 1, text: "pasted one" },
            { kind: "cellText", sheetId, col: 0, row: 2, text: "pasted two" },
        ]);

        // Sam, meanwhile, types in the column beside it.
        const theirs = run(base, sam, [
            { kind: "cellText", sheetId, col: 1, row: 0, text: "sam typed" },
        ]);

        const onMine = merge(mine, theirs).doc;
        const onTheirs = merge(theirs, mine).doc;

        expect(column(onMine, 0)).toEqual([
            "perm",
            "pasted one",
            "pasted two",
            "cap bad",
            "extend",
        ]);
        expect(column(onTheirs, 0)).toEqual(column(onMine, 0));
        expect(column(onTheirs, 1)).toEqual(column(onMine, 1));
        expect(liveCells(onMine.sheets[sheetId], 0)).toHaveLength(5);
    });

    it("does not double the column when both sides paste at once", () => {
        const alex = ctxFor("alex", 1_000);
        const sam = ctxFor("sam", 5_000);
        const mine = run(base, alex, openSpanOps(sheetId, 0, 1, 1));
        const theirs = run(base, sam, openSpanOps(sheetId, 0, 1, 1));

        const onMine = merge(mine, theirs).doc;
        const onTheirs = merge(theirs, mine).doc;

        // Two concurrent opens are two rows, on both machines. Not four, and
        // not a different number on each.
        expect(column(onMine, 0)).toEqual(column(onTheirs, 0));
        expect(column(onMine, 0)).toEqual(["perm", null, null, "cap bad", "extend"]);
    });
});

describe("the row growth a paste causes", () => {
    // The paste says what it opened, in the columns it landed in. Taking
    // Handsontable's row growth as well would open the same rows again, this
    // time across every column of the sheet.
    it("is not taken as an op of its own", () => {
        for (const source of ["auto", "populateFromArray", "CopyPaste.paste"]) {
            expect(rowOpFromHook("insert", sheetId, 2, 3, source)).toEqual([]);
        }
    });

    it("is still taken for a row the debater inserted", () => {
        expect(rowOpFromHook("insert", sheetId, 2, 1, "ContextMenu.rowAbove")).toEqual([
            { kind: "insertRow", sheetId, row: 2 },
        ]);
    });
});
