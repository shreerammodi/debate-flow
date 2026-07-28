import ExcelJS from "exceljs";
import { describe, expect, it } from "vitest";

import type { Contacts } from "@/lib/collab/contacts";
import { applyInfoWorksheet, maybeAddRfdWorksheet } from "@/lib/export/infoSheet";
import { makeFlowRound } from "@/lib/model/flow";

const roundWith = (patch: object) => {
    const round = makeFlowRound({});
    Object.assign(round.scouting, patch);
    return round;
};

describe("applyInfoWorksheet", () => {
    it("writes scouting fields into the Info worksheet", () => {
        const wb = new ExcelJS.Workbook();
        const round = roundWith({
            tournament: "TOC",
            judge: "Judge Judy",
            affSchool: "Alpha HS",
            aff: {
                first: { first: "Ada", last: "L" },
                second: { first: "Ben", last: "M" },
            },
        });
        applyInfoWorksheet(wb, round);
        const ws = wb.getWorksheet("Info")!;
        expect(ws.getCell("B2").value).toBe("TOC");
        expect(ws.getCell("B5").value).toBe("Judge Judy");
        expect(ws.getCell("B7").value).toBe("Alpha HS");
        expect(ws.getCell("B8").value).toBe("Ada L");
        expect(ws.getCell("B9").value).toBe("Ben M");
    });

    it("labels the sides and speaker slots the way the event names them", () => {
        const wb = new ExcelJS.Workbook();
        const round = makeFlowRound({ event: "parli" });
        round.scouting.decision = { vote: "neg" };
        applyInfoWorksheet(wb, round);
        const ws = wb.getWorksheet("Info")!;
        expect([7, 8, 9, 11, 12, 13].map((r) => ws.getCell(r, 1).value)).toEqual([
            "Gov School",
            "PM",
            "MG",
            "Opp School",
            "LO",
            "MO",
        ]);
        expect(ws.getCell("B15").value).toBe("OPP");
    });
});

describe("maybeAddRfdWorksheet", () => {
    it("adds an RFD worksheet when a decision exists", () => {
        const wb = new ExcelJS.Workbook();
        const round = roundWith({ decision: { vote: "neg", rfd: "Dropped the disad." } });
        maybeAddRfdWorksheet(wb, round);
        const ws = wb.getWorksheet("RFD")!;
        expect(ws.getCell("B1").value).toBe("NEG");
        expect(ws.getCell("A2").value).toBe("Dropped the disad.");
    });

    it("skips the worksheet when there is no vote and no rfd", () => {
        const wb = new ExcelJS.Workbook();
        maybeAddRfdWorksheet(wb, makeFlowRound({}));
        expect(wb.getWorksheet("RFD")).toBeUndefined();
    });

    const RAE = "aaa11111aaa";
    const SAM = "bbb22222bbb";
    const contacts: Contacts = {
        [RAE]: { name: "Rae", role: "coach" },
        [SAM]: { name: "Sam", role: "partner" },
    };

    it("adds the worksheet for peer notes alone, with no vote and no local rfd", () => {
        const wb = new ExcelJS.Workbook();
        const round = roundWith({ decision: { peerNotes: { [RAE]: "aff on T" } } });
        maybeAddRfdWorksheet(wb, round, contacts);
        const ws = wb.getWorksheet("RFD")!;
        expect(ws.getCell("A1").value).toBe("Decision");
        expect(ws.getCell("B1").value).toBeNull();
        expect(ws.getCell("A2").value).toBeNull();
        expect(ws.getCell("A3").value).toBe("Notes from Rae");
        expect(ws.getCell("A4").value).toBe("aff on T");
    });

    it("writes each peer under a label, after the local notes", () => {
        const wb = new ExcelJS.Workbook();
        const round = roundWith({
            decision: {
                vote: "aff",
                rfd: "my own voter",
                peerNotes: { [SAM]: "neg on case", [RAE]: "aff on T" },
            },
        });
        maybeAddRfdWorksheet(wb, round, contacts);
        const ws = wb.getWorksheet("RFD")!;
        expect(ws.getCell("B1").value).toBe("AFF");
        expect(ws.getCell("A2").value).toBe("my own voter");
        expect(ws.getCell("A4").value).toBe("Notes from Rae");
        expect(ws.getCell("A5").value).toBe("aff on T");
        expect(ws.getCell("A7").value).toBe("Notes from Sam");
        expect(ws.getCell("A8").value).toBe("neg on case");
    });

    it("labels an unknown peer with the short form of its EndpointId", () => {
        const wb = new ExcelJS.Workbook();
        const round = roundWith({ decision: { peerNotes: { "0123456789abcdef": "dropped it" } } });
        maybeAddRfdWorksheet(wb, round, contacts);
        expect(wb.getWorksheet("RFD")!.getCell("A3").value).toBe("Notes from 01234567");
    });

    it("wraps every author's notes so long text stays readable", () => {
        const wb = new ExcelJS.Workbook();
        const round = roundWith({
            decision: { rfd: "my own voter", peerNotes: { [RAE]: "aff on T" } },
        });
        maybeAddRfdWorksheet(wb, round, contacts);
        const ws = wb.getWorksheet("RFD")!;
        expect(ws.getCell("A2").alignment).toMatchObject({ wrapText: true, vertical: "top" });
        expect(ws.getCell("A5").alignment).toMatchObject({ wrapText: true, vertical: "top" });
    });

    it("skips the worksheet when every peer note is blank", () => {
        const wb = new ExcelJS.Workbook();
        const round = roundWith({ decision: { peerNotes: { [RAE]: "   " } } });
        maybeAddRfdWorksheet(wb, round, contacts);
        expect(wb.getWorksheet("RFD")).toBeUndefined();
    });
});
