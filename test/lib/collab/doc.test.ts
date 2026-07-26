import { describe, expect, it } from "vitest";

import { liveCells, projectDoc, seedDoc, sheetWidth } from "@/lib/collab/doc";
import { ORIGIN_STAMP } from "@/lib/collab/stamp";
import { cellKey } from "@/lib/collab/types";
import { makeFlowRound, type FlowRound } from "@/lib/model/flow";

function roundWithData(): FlowRound {
    const round = makeFlowRound({});
    const flow = round.sheets.find((s) => s.kind !== "cx")!;
    flow.data = [
        ["perm do both", "no link"],
        ["cap bad", null],
    ];
    flow.meta = { "0,0": { bold: true }, "1,0": { card: true } };
    return round;
}

describe("seedDoc", () => {
    it("derives an identical replica on two machines from one file", () => {
        const round = roundWithData();
        expect(seedDoc(round)).toEqual(seedDoc(structuredClone(round)));
    });

    it("stamps every seeded value with the origin, below any real write", () => {
        const doc = seedDoc(roundWithData());
        const sheet = Object.values(doc.sheets).find((s) => s.fields.kind.value !== "cx")!;
        expect(Object.values(sheet.cells).every((c) => c.textStamp === ORIGIN_STAMP)).toBe(true);
        expect(doc.round.event.stamp).toEqual(ORIGIN_STAMP);
    });

    it("credits a seeded cell to no actor, so both peers agree on its identity", () => {
        const doc = seedDoc(roundWithData());
        const sheet = Object.values(doc.sheets).find((s) => s.fields.kind.value !== "cx")!;
        expect(Object.keys(sheet.cells)).toContain(cellKey(0, liveCells(sheet, 0)[0].rank, ""));
        expect(Object.values(sheet.cells).every((c) => c.actor === "")).toBe(true);
    });

    it("carries the round's scalar leaves", () => {
        const round = makeFlowRound({ event: "pf", firstSide: "neg" });
        round.scouting.tournament = "Harvard";
        const doc = seedDoc(round);
        expect(doc.roundId).toBe(round.id);
        expect(doc.round.event.value).toBe("pf");
        expect(doc.round.firstSide.value).toBe("neg");
        expect(doc.round["scouting.tournament"].value).toBe("Harvard");
    });

    it("pads a ragged sheet to the rectangle the grid shows", () => {
        const doc = seedDoc(roundWithData());
        const sheet = Object.values(doc.sheets).find((s) => s.fields.kind.value !== "cx")!;
        expect(sheetWidth(sheet)).toBe(2);
        expect(liveCells(sheet, 1)).toHaveLength(2);
        expect(liveCells(sheet, 1)[1].text).toBeNull();
    });
});

describe("projectDoc", () => {
    it("round-trips a round through the replica", () => {
        const round = roundWithData();
        expect(projectDoc(seedDoc(round), round)).toEqual(round);
    });

    it("takes createdAt and updatedAt from the local round, never from a peer", () => {
        const round = roundWithData();
        const doc = seedDoc(round);
        const later = { ...round, updatedAt: round.updatedAt + 5_000 };
        expect(projectDoc(doc, later).updatedAt).toBe(round.updatedAt + 5_000);
    });

    it("drops a deleted sheet and a deleted cell", () => {
        const round = roundWithData();
        const doc = seedDoc(round);
        const sheet = Object.values(doc.sheets).find((s) => s.fields.kind.value !== "cx")!;
        const first = liveCells(sheet, 0)[0];
        sheet.cells[cellKey(0, first.rank, "")].deleted = { ms: 1, counter: 0, actor: "sam" };
        const projected = projectDoc(doc, round);
        const flow = projected.sheets.find((s) => s.kind !== "cx")!;
        expect(flow.data).toEqual([
            ["cap bad", "no link"],
            [null, null],
        ]);
        expect(flow.meta).toEqual({ "0,0": { card: true } });

        Object.values(doc.sheets).forEach((s) => {
            if (s.fields.kind.value !== "cx") s.deleted = { ms: 2, counter: 0, actor: "sam" };
        });
        expect(projectDoc(doc, round).sheets.map((s) => s.kind)).toEqual(["cx"]);
    });

    it("orders sheets by order then id, identically on every peer", () => {
        const round = makeFlowRound({});
        const doc = seedDoc(round);
        const ids = Object.keys(doc.sheets);
        for (const id of ids) doc.sheets[id].fields.order = { value: 0, stamp: ORIGIN_STAMP };
        const expected = ids.slice().sort();
        expect(projectDoc(doc, round).sheets.map((s) => s.id)).toEqual(expected);
    });
});
