import { describe, expect, it } from "vitest";

import { EVENTS } from "@/lib/format/events";
import {
    compareSheets,
    firstFlowSheetId,
    makeCxFlowSheet,
    makeFlowRound,
    makeFlowSheet,
    normalizeFlow,
    sortedSheets,
    type FlowRound,
} from "@/lib/model/flow";

describe("makeFlowRound", () => {
    it("creates a CX sheet plus one flow sheet grouped to the first speaker", () => {
        const r = makeFlowRound();
        expect(r.sheets).toHaveLength(2);
        const cx = r.sheets.find((s) => s.kind === "cx")!;
        const flow = r.sheets.find((s) => s.kind !== "cx")!;
        expect(cx.order).toBe(-1);
        expect(flow.title).toBe("1.");
        expect(flow.group).toBe("aff");
        expect(flow.startSpeechId).toBeUndefined();
        expect(flow.data).toEqual([]);
        expect(flow.meta).toEqual({});
    });

    it("groups the first sheet to neg when neg speaks first", () => {
        const r = makeFlowRound({ event: "pf", firstSide: "neg" });
        const flow = r.sheets.find((s) => s.kind !== "cx")!;
        expect(flow.group).toBe("neg");
        expect(flow.startSpeechId).toBeUndefined();
    });
});

describe("multi-event round fields", () => {
    it("makeFlowRound defaults to a policy round with aff speaking first", () => {
        const round = makeFlowRound({});
        expect(round.event).toBe("policy");
        expect(round.firstSide).toBe("aff");
        expect(round.sheets.find((s) => s.kind === "cx")?.title).toBe("CX");
    });

    it("makeFlowRound builds a pf round with the event's cross-ex title", () => {
        const round = makeFlowRound({ event: "pf", firstSide: "neg" });
        expect(round.event).toBe("pf");
        expect(round.firstSide).toBe("neg");
        expect(round.sheets.find((s) => s.kind === "cx")?.title).toBe(EVENTS.pf.crossEx?.title);
    });

    it("gives a parli round no cross-ex sheet, and never backfills one", () => {
        const round = makeFlowRound({ event: "parli" });
        expect(round.sheets).toHaveLength(1);
        expect(round.sheets[0].kind).toBe("flow");
        expect(normalizeFlow(round).sheets.filter((s) => s.kind === "cx")).toHaveLength(0);
    });

    it("normalizeFlow backfills event and firstSide on legacy rounds", () => {
        const legacy = makeFlowRound({});
        delete (legacy as Partial<FlowRound>).event;
        delete (legacy as Partial<FlowRound>).firstSide;
        const normalized = normalizeFlow(legacy);
        expect(normalized.event).toBe("policy");
        expect(normalized.firstSide).toBe("aff");
    });

    it("makeFlowSheet leaves startSpeechId unset", () => {
        const sheet = makeFlowSheet({ title: "1.", group: "neg", order: 0 });
        expect(sheet.startSpeechId).toBeUndefined();
    });
});

describe("normalizeFlow", () => {
    it("fills defaults and guarantees exactly one CX sheet", () => {
        const raw = {
            id: "r1",
            createdAt: 1,
            updatedAt: 2,
            scouting: undefined,
            sheets: [{ id: "s1", title: "Aff", group: "aff", order: 0 }],
        } as unknown as FlowRound;
        const r = normalizeFlow(raw);
        expect(r.scouting.aff.first).toEqual({ first: "", last: "" });
        expect(r.sheets.filter((s) => s.kind === "cx")).toHaveLength(1);
        const s1 = r.sheets.find((s) => s.id === "s1")!;
        expect(s1.kind).toBe("flow");
        expect(s1.data).toEqual([]);
        expect(s1.meta).toEqual({});
    });

    it("does not duplicate an existing CX sheet and does not mutate input", () => {
        const raw = makeFlowRound({});
        const before = JSON.parse(JSON.stringify(raw));
        const r = normalizeFlow(raw);
        expect(r.sheets.filter((s) => s.kind === "cx")).toHaveLength(1);
        expect(JSON.parse(JSON.stringify(raw))).toEqual(before);
    });
});

describe("sheet ordering", () => {
    it("sortedSheets sorts by order; firstFlowSheetId skips CX", () => {
        const r = makeFlowRound({});
        const extra = makeFlowSheet({ title: "DA", group: "neg", order: 5 });
        const round = { ...r, sheets: [extra, ...r.sheets] };
        expect(sortedSheets(round).map((s) => s.title)).toEqual(["CX", "1.", "DA"]);
        expect(firstFlowSheetId(round)).toBe(r.sheets.find((s) => s.kind !== "cx")!.id);
        expect(firstFlowSheetId({ ...round, sheets: [makeCxFlowSheet()] })).not.toBeNull();
    });

    it("breaks an order tie on the sheet id, so two peers agree", () => {
        const a = { ...makeFlowSheet({ title: "A", group: "aff", order: 2 }), id: "sheet-b" };
        const b = { ...makeFlowSheet({ title: "B", group: "neg", order: 2 }), id: "sheet-a" };
        expect(compareSheets(a, b)).toBeGreaterThan(0);
        expect(sortedSheets({ ...makeFlowRound({}), sheets: [a, b] }).map((s) => s.id)).toEqual([
            "sheet-a",
            "sheet-b",
        ]);
        expect(sortedSheets({ ...makeFlowRound({}), sheets: [b, a] }).map((s) => s.id)).toEqual([
            "sheet-a",
            "sheet-b",
        ]);
    });

    it("firstFlowSheetId is null for an empty sheet list", () => {
        const r = { ...makeFlowRound({}), sheets: [] };
        expect(firstFlowSheetId(r)).toBeNull();
    });
});
