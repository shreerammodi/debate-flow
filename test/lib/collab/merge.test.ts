import { describe, expect, it } from "vitest";

import { liveCells, projectDoc, seedDoc } from "@/lib/collab/doc";
import { merge } from "@/lib/collab/merge";
import { applyOp, type OpContext } from "@/lib/collab/ops";
import { createClock } from "@/lib/collab/stamp";
import type { CollabDoc } from "@/lib/collab/types";
import { makeFlowRound, type FlowRound } from "@/lib/model/flow";

function peer(actor: string, start: number): OpContext {
    let t = start;
    return { actor, clock: createClock(actor, () => t++) };
}

function base(): { round: FlowRound; doc: CollabDoc; sheetId: string } {
    const round = makeFlowRound({});
    const flow = round.sheets.find((s) => s.kind !== "cx")!;
    flow.data = [
        ["perm", "link"],
        ["cap bad", "turn"],
    ];
    return { round, doc: seedDoc(round), sheetId: flow.id };
}

describe("merge", () => {
    it("keeps the later write when two peers edit different cells", () => {
        const { round, doc, sheetId } = base();
        const a = applyOp(
            doc,
            { kind: "cellText", sheetId, col: 0, row: 0, text: "A" },
            peer("alex", 10),
        );
        const b = applyOp(
            doc,
            { kind: "cellText", sheetId, col: 1, row: 0, text: "B" },
            peer("sam", 20),
        );
        const merged = merge(a, b).doc;
        const data = projectDoc(merged, round).sheets.find((s) => s.id === sheetId)!.data;
        expect(data[0]).toEqual(["A", "B"]);
    });

    it("resolves the same cell by wall time, not by arrival", () => {
        const { round, doc, sheetId } = base();
        const early = applyOp(
            doc,
            { kind: "cellText", sheetId, col: 0, row: 0, text: "early" },
            peer("alex", 10),
        );
        const late = applyOp(
            doc,
            { kind: "cellText", sheetId, col: 0, row: 0, text: "late" },
            peer("sam", 99),
        );
        const read = (r: FlowRound) => r.sheets.find((s) => s.id === sheetId)!.data[0][0];
        expect(read(projectDoc(merge(early, late).doc, round))).toBe("late");
        expect(read(projectDoc(merge(late, early).doc, round))).toBe("late");
    });

    it("keeps text and meta apart, so a bold toggle never reverts a partner's text", () => {
        const { round, doc, sheetId } = base();
        const typed = applyOp(
            doc,
            { kind: "cellText", sheetId, col: 0, row: 0, text: "typed" },
            peer("alex", 50),
        );
        const bolded = applyOp(
            doc,
            { kind: "cellMeta", sheetId, col: 0, row: 0, meta: { bold: true } },
            peer("sam", 90),
        );
        const sheet = projectDoc(merge(typed, bolded).doc, round).sheets.find(
            (s) => s.id === sheetId,
        )!;
        expect(sheet.data[0][0]).toBe("typed");
        expect(sheet.meta["0,0"]).toEqual({ bold: true });
    });

    it("lets a delete win over a later write", () => {
        const { round, doc, sheetId } = base();
        const deleted = applyOp(doc, { kind: "removeRow", sheetId, row: 0 }, peer("alex", 10));
        const written = applyOp(
            doc,
            { kind: "cellText", sheetId, col: 0, row: 0, text: "late" },
            peer("sam", 900),
        );
        const data = projectDoc(merge(deleted, written).doc, round).sheets.find(
            (s) => s.id === sheetId,
        )!.data;
        expect(data).toEqual([["cap bad", "turn"]]);
    });

    it("reports every non-empty cell a delete discards", () => {
        const { doc, sheetId } = base();
        const written = applyOp(
            doc,
            { kind: "cellText", sheetId, col: 0, row: 0, text: "keep me" },
            peer("sam", 900),
        );
        const deleted = applyOp(doc, { kind: "removeRow", sheetId, row: 0 }, peer("alex", 10));
        const { dropped } = merge(written, deleted);
        expect(dropped.map((d) => ({ text: d.text, col: d.col, by: d.deletedBy }))).toEqual([
            { text: "keep me", col: 0, by: "alex" },
            { text: "link", col: 1, by: "alex" },
        ]);
        expect(dropped.every((d) => d.sheetId === sheetId)).toBe(true);
    });

    it("reports nothing when the merge changes nothing", () => {
        const { doc, sheetId } = base();
        const written = applyOp(
            doc,
            { kind: "cellText", sheetId, col: 0, row: 0, text: "keep me" },
            peer("sam", 900),
        );
        const deleted = applyOp(doc, { kind: "removeRow", sheetId, row: 0 }, peer("alex", 10));
        const once = merge(written, deleted);
        expect(merge(once.doc, deleted).dropped).toEqual([]);
    });

    it("never resurrects a cell a peer still holds", () => {
        const { round, doc, sheetId } = base();
        const deleted = applyOp(doc, { kind: "removeRow", sheetId, row: 0 }, peer("alex", 10));
        const data = projectDoc(merge(deleted, doc).doc, round).sheets.find(
            (s) => s.id === sheetId,
        )!.data;
        expect(data).toEqual([["cap bad", "turn"]]);
    });

    it("keeps both cells when two peers insert at one row", () => {
        const { doc, sheetId } = base();
        const a = applyOp(doc, { kind: "insertCell", sheetId, col: 0, row: 1 }, peer("alex", 10));
        const b = applyOp(doc, { kind: "insertCell", sheetId, col: 0, row: 1 }, peer("sam", 10));
        expect(liveCells(merge(a, b).doc.sheets[sheetId], 0)).toHaveLength(4);
        expect(liveCells(merge(b, a).doc.sheets[sheetId], 0).map((c) => c.actor)).toEqual([
            "",
            "alex",
            "sam",
            "",
        ]);
    });

    it("takes the first delete when two peers delete one sheet", () => {
        const { doc, sheetId } = base();
        const a = applyOp(doc, { kind: "removeSheet", sheetId }, peer("alex", 10));
        const b = applyOp(doc, { kind: "removeSheet", sheetId }, peer("sam", 40));
        expect(merge(a, b).doc.sheets[sheetId].deleted).toEqual(
            merge(b, a).doc.sheets[sheetId].deleted,
        );
        expect(merge(a, b).doc.sheets[sheetId].deleted?.actor).toBe("alex");
    });

    it("adopts a sheet the local replica has never seen", () => {
        const { round, doc, sheetId } = base();
        const remote = applyOp(
            doc,
            {
                kind: "addSheet",
                sheet: {
                    id: "sheet-remote",
                    title: "DA",
                    group: "neg",
                    order: 3,
                    kind: "flow",
                    data: [["shell"]],
                    meta: {},
                },
            },
            peer("sam", 10),
        );
        const merged = merge(doc, remote).doc;
        expect(projectDoc(merged, round).sheets.map((s) => s.title)).toContain("DA");
        expect(merged.sheets[sheetId]).toBeDefined();
    });
});
