import { describe, expect, it } from "vitest";

import { deltaSince, isEmptyDelta, vectorOf } from "@/lib/collab/delta";
import { seedDoc } from "@/lib/collab/doc";
import { merge } from "@/lib/collab/merge";
import { applyOp, type OpContext } from "@/lib/collab/ops";
import { createClock } from "@/lib/collab/stamp";
import { makeFlowRound, type FlowRound } from "@/lib/model/flow";

function peer(actor: string, start: number): OpContext {
    let t = start;
    return { actor, clock: createClock(actor, () => t++) };
}

function base(): { round: FlowRound; sheetId: string } {
    const round = makeFlowRound({});
    const flow = round.sheets.find((s) => s.kind !== "cx")!;
    flow.data = [
        ["perm", "link"],
        ["cap bad", "turn"],
    ];
    return { round, sheetId: flow.id };
}

describe("vectorOf", () => {
    it("reports one origin entry for a freshly seeded document", () => {
        const { round } = base();
        expect(vectorOf(seedDoc(round))).toEqual({ "": { ms: 0, counter: 0, actor: "" } });
    });

    it("reports the highest stamp each actor wrote", () => {
        const { round, sheetId } = base();
        let doc = seedDoc(round);
        const alex = peer("alex", 10);
        doc = applyOp(doc, { kind: "cellText", sheetId, col: 0, row: 0, text: "a" }, alex);
        doc = applyOp(doc, { kind: "cellText", sheetId, col: 0, row: 1, text: "b" }, alex);
        expect(vectorOf(doc).alex.ms).toBe(11);
    });
});

describe("deltaSince", () => {
    it("carries nothing when the far side has seen everything", () => {
        const { round } = base();
        const doc = seedDoc(round);
        expect(isEmptyDelta(deltaSince(doc, vectorOf(doc)))).toBe(true);
    });

    it("carries only the cell that changed", () => {
        const { round, sheetId } = base();
        const before = seedDoc(round);
        const seen = vectorOf(before);
        const after = applyOp(
            before,
            { kind: "cellText", sheetId, col: 1, row: 0, text: "new" },
            peer("alex", 10),
        );
        const delta = deltaSince(after, seen);
        expect(Object.keys(delta.sheets)).toEqual([sheetId]);
        const cells = Object.values(delta.sheets[sheetId].cells);
        expect(cells).toHaveLength(1);
        expect(cells[0].text).toBe("new");
    });

    it("carries a round field that changed and no sheet at all", () => {
        const { round } = base();
        const before = seedDoc(round);
        const after = applyOp(
            before,
            { kind: "roundField", path: "scouting.judge", value: "Ito" },
            peer("alex", 10),
        );
        const delta = deltaSince(after, vectorOf(before));
        expect(Object.keys(delta.round)).toEqual(["scouting.judge"]);
        expect(Object.keys(delta.sheets)).toEqual([]);
    });

    it("carries a tombstone the far side has not seen", () => {
        const { round, sheetId } = base();
        const before = seedDoc(round);
        const after = applyOp(before, { kind: "removeRow", sheetId, row: 0 }, peer("alex", 10));
        const delta = deltaSince(after, vectorOf(before));
        const cells = Object.values(delta.sheets[sheetId].cells);
        expect(cells).toHaveLength(2);
        expect(cells.every((c) => c.deleted !== null)).toBe(true);
    });

    it("carries a sheet field without dragging every cell of the sheet", () => {
        const { round, sheetId } = base();
        const before = seedDoc(round);
        const after = applyOp(
            before,
            { kind: "sheetField", sheetId, path: "title", value: "T" },
            peer("alex", 10),
        );
        const delta = deltaSince(after, vectorOf(before));
        expect(Object.keys(delta.sheets[sheetId].fields)).toEqual(["title"]);
        expect(Object.keys(delta.sheets[sheetId].cells)).toEqual([]);
    });

    it("brings the far side exactly to where the near side is", () => {
        const { round, sheetId } = base();
        const far = seedDoc(round);
        let near = far;
        const alex = peer("alex", 10);
        near = applyOp(near, { kind: "cellText", sheetId, col: 0, row: 0, text: "one" }, alex);
        near = applyOp(near, { kind: "insertRow", sheetId, row: 1 }, alex);
        near = applyOp(near, { kind: "cellText", sheetId, col: 1, row: 1, text: "two" }, alex);

        expect(merge(far, deltaSince(near, vectorOf(far))).doc).toEqual(near);
    });

    it("answers a repair for a peer that missed the middle of a burst", () => {
        const { round, sheetId } = base();
        const start = seedDoc(round);
        const alex = peer("alex", 10);
        const one = applyOp(start, { kind: "cellText", sheetId, col: 0, row: 0, text: "1" }, alex);
        const two = applyOp(one, { kind: "cellText", sheetId, col: 0, row: 1, text: "2" }, alex);
        const three = applyOp(two, { kind: "cellText", sheetId, col: 1, row: 0, text: "3" }, alex);

        // The far side got the first write and then the link dropped.
        const far = merge(start, deltaSince(one, vectorOf(start))).doc;
        expect(merge(far, deltaSince(three, vectorOf(far))).doc).toEqual(three);
    });

    it("costs nothing between two peers that opened the same file", () => {
        const { round } = base();
        // Both seeded from one file, so every value shares the origin stamp.
        expect(isEmptyDelta(deltaSince(seedDoc(round), vectorOf(seedDoc(round))))).toBe(true);
    });
});
