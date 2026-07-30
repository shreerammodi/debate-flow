import { describe, expect, it } from "vitest";

import { liveCells, projectDoc, projectSheet, seedDoc, sheetWidth } from "@/lib/collab/doc";
import { applyOp, type OpContext } from "@/lib/collab/ops";
import { createClock, ORIGIN_STAMP } from "@/lib/collab/stamp";
import { cellKey } from "@/lib/collab/types";
import { compareSheets, makeFlowRound, makeFlowSheet, type FlowRound } from "@/lib/model/flow";
import { MAX_ROUND_CELLS, paddedCells } from "@/lib/persistence/flowFile";

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

/**
 * A merge builds a new object only for the sheets it touched, so the sheets it
 * did not are handed back as they were. One cell arriving from a partner
 * otherwise re-derives every sheet in the round, thirty times a second while
 * they type.
 */
describe("projecting a round a partner just changed", () => {
    it("hands back the very same object for a sheet the merge left alone", () => {
        const round = roundWithData();
        const settled = seedDoc(round);
        const base = projectDoc(settled, round);

        const sam: OpContext = { actor: "sam", clock: createClock("sam", () => 5_000) };
        const touchedId = round.sheets[1].id;
        const after = applyOp(
            settled,
            { kind: "cellText", sheetId: touchedId, col: 0, row: 0, text: "theirs" },
            sam,
        );

        const next = projectDoc(after, base, settled);
        for (const sheet of next.sheets) {
            const was = base.sheets.find((s) => s.id === sheet.id)!;
            if (sheet.id === touchedId) expect(sheet).not.toBe(was);
            else expect(sheet, `${sheet.id} was re-derived for nothing`).toBe(was);
        }
    });

    it("describes the same round either way", () => {
        const round = roundWithData();
        const settled = seedDoc(round);
        const base = projectDoc(settled, round);
        const sam: OpContext = { actor: "sam", clock: createClock("sam", () => 5_000) };
        const after = applyOp(
            settled,
            { kind: "cellText", sheetId: round.sheets[1].id, col: 1, row: 1, text: "theirs" },
            sam,
        );
        expect(projectDoc(after, base, settled)).toEqual(projectDoc(after, base));
    });

    it("re-derives everything when the caller has nothing settled to offer", () => {
        const round = roundWithData();
        const doc = seedDoc(round);
        const base = projectDoc(doc, round);
        for (const sheet of projectDoc(doc, base).sheets) {
            expect(sheet).not.toBe(base.sheets.find((s) => s.id === sheet.id));
        }
    });
});

/**
 * The round's cells are bounded across sheets, because the file counts them
 * across sheets. What matters as much as the bound is that it is unreachable
 * from anything a debater can type.
 */
describe("the round's cell budget", () => {
    /** A fat but ordinary elim: six sheets, a few hundred rows, eight speeches. */
    function realisticRound(): FlowRound {
        const round = makeFlowRound({});
        while (round.sheets.length < 6) {
            const order = round.sheets.length;
            round.sheets.push(makeFlowSheet({ title: `${order}.`, group: "neg", order }));
        }
        for (const sheet of round.sheets) {
            sheet.data = Array.from({ length: 220 }, (_, r) =>
                Array.from({ length: 8 }, (_, c) => (r % 4 === 0 ? null : `arg ${r}.${c}`)),
            );
            sheet.meta = { "0,1": { bold: true }, "17,3": { card: true } };
        }
        return round;
    }

    it("is inert on an ordinary round, which projects byte for byte as it did before", () => {
        const round = realisticRound();
        const doc = seedDoc(round);
        // `projectSheet` with no room given is the projection as it stood before
        // the budget: one sheet cannot reach MAX_ROUND_CELLS on its own, MAX_COL
        // times MAX_ROWS being under it. So this is the old output sheet for
        // sheet, and the budgeted round has to serialize to the same bytes.
        const unbudgeted = Object.values(doc.sheets)
            .map((sheet) => projectSheet(sheet))
            .sort(compareSheets);
        expect(JSON.stringify(projectDoc(doc, round).sheets)).toBe(JSON.stringify(unbudgeted));

        // Including through the reuse path, which the budget must not cost: a
        // settled sheet is still the very same object, not derived again.
        const base = projectDoc(doc, round);
        const reused = projectDoc(doc, base, doc);
        expect(JSON.stringify(reused.sheets)).toBe(JSON.stringify(unbudgeted));
        expect(reused.sheets.every((s, i) => s === base.sheets[i])).toBe(true);

        // The margin: no sheet here is within a factor of two of the smallest
        // share the budget could offer it, which is one 512th of the round.
        for (const sheet of unbudgeted) {
            expect(paddedCells(sheet.data)).toBeLessThan(MAX_ROUND_CELLS / 512);
        }
    });
});
