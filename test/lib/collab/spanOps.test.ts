import { beforeEach, describe, expect, it } from "vitest";

import { projectDoc, seedDoc } from "@/lib/collab/doc";
import { applyOp, type OpContext } from "@/lib/collab/ops";
import { replaceSpanOps } from "@/lib/collab/spanOps";
import { createClock } from "@/lib/collab/stamp";
import type { CollabDoc } from "@/lib/collab/types";
import { modelCol } from "@/lib/grid/colSpace";
import { makeFlowRound, type FlowRound } from "@/lib/model/flow";

let round: FlowRound;
let sheetId: string;
let doc: CollabDoc;
let alex: OpContext;

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
    let t = 1_000;
    alex = { actor: "alex", clock: createClock("alex", () => t++) };
});

function run(ops: ReturnType<typeof replaceSpanOps>): (string | null)[] {
    let next = doc;
    for (const op of ops) next = applyOp(next, op, alex);
    return projectDoc(next, round)
        .sheets.find((s) => s.id === sheetId)!
        .data.map((r) => r[0]);
}

describe("replaceSpanOps", () => {
    it("rewrites a span in place", () => {
        const ops = replaceSpanOps(sheetId, modelCol(0), 1, ["X", "Y"]);
        expect(run(ops)).toEqual(["a0", "X", "Y", "a3"]);
    });

    it("expresses a block move down as ops, with no re-seed", () => {
        // "a1" moves below "a2": the whole span is rewritten in its new order.
        const ops = replaceSpanOps(sheetId, modelCol(0), 1, ["a2", "a1"]);
        expect(run(ops)).toEqual(["a0", "a2", "a1", "a3"]);
    });

    it("leaves neighbouring columns alone", () => {
        let next = doc;
        for (const op of replaceSpanOps(sheetId, modelCol(0), 0, ["Z"])) {
            next = applyOp(next, op, alex);
        }
        const data = projectDoc(next, round).sheets.find((s) => s.id === sheetId)!.data;
        expect(data.map((r) => r[1])).toEqual(["b0", "b1", "b2", "b3"]);
    });

    it("uses only ops a peer can apply, never a re-seed", () => {
        const kinds = new Set(
            replaceSpanOps(sheetId, modelCol(0), 1, ["X", "Y"]).map((o) => o.kind),
        );
        expect([...kinds].sort()).toEqual(["cellText", "insertCell", "removeCell"]);
    });

    it("keeps the column's height unchanged", () => {
        expect(run(replaceSpanOps(sheetId, modelCol(0), 0, ["P", "Q", "R", "S"]))).toHaveLength(4);
    });

    it("does nothing for an empty span", () => {
        expect(replaceSpanOps(sheetId, modelCol(0), 1, [])).toEqual([]);
    });
});
