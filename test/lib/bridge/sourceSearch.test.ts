import { describe, expect, it } from "vitest";

import { findCellsBySourceKey } from "@/lib/bridge/sourceSearch";
import type { CellMeta, CellSource, FlowRound, FlowSheet } from "@/lib/model/flow";

const src = (key: string): CellSource => ({ app: "cardmirror", token: `t-${key}`, key });

function sheet(
    id: string,
    title: string,
    order: number,
    meta: Record<string, CellMeta>,
): FlowSheet {
    return { id, title, group: "aff", order, kind: "flow", data: [], meta };
}

function round(sheets: FlowSheet[]): FlowRound {
    return {
        id: "round1",
        createdAt: 0,
        updatedAt: 0,
        event: "policy",
        scouting: {
            aff: {
                first: { first: "", last: "" },
                second: { first: "", last: "" },
            },
            neg: {
                first: { first: "", last: "" },
                second: { first: "", last: "" },
            },
        },
        sheets,
    };
}

describe("findCellsBySourceKey", () => {
    it("returns hits in sheet, then row, then column order", () => {
        // Sheets and meta keys are both out of order, so only the sort can
        // produce the expected sequence.
        const r = round([
            sheet("s2", "1AR", 1, {
                "0,0": { source: src("k1") },
            }),
            sheet("s1", "2AC", 0, {
                "5,1": { source: src("k2") },
                "2,3": { source: src("k1") },
                "2,0": { source: src("k2") },
            }),
        ]);

        expect(findCellsBySourceKey(r, ["k1", "k2"])).toEqual([
            { sheetId: "s1", sheetTitle: "2AC", row: 2, col: 0 },
            { sheetId: "s1", sheetTitle: "2AC", row: 2, col: 3 },
            { sheetId: "s1", sheetTitle: "2AC", row: 5, col: 1 },
            { sheetId: "s2", sheetTitle: "1AR", row: 0, col: 0 },
        ]);
    });

    it("skips cells whose provenance key is not wanted", () => {
        const r = round([
            sheet("s1", "2AC", 0, {
                "0,0": { source: src("k1") },
                "1,0": { source: src("other") },
                "2,0": { bold: true },
            }),
        ]);

        expect(findCellsBySourceKey(r, ["k1"])).toEqual([
            { sheetId: "s1", sheetTitle: "2AC", row: 0, col: 0 },
        ]);
    });

    it("returns nothing when no cell matches", () => {
        const r = round([sheet("s1", "2AC", 0, { "0,0": { source: src("k1") } })]);

        expect(findCellsBySourceKey(r, ["nope"])).toEqual([]);
    });

    it("returns nothing for an empty key list", () => {
        const r = round([sheet("s1", "2AC", 0, { "0,0": { source: src("k1") } })]);

        expect(findCellsBySourceKey(r, [])).toEqual([]);
    });
});
