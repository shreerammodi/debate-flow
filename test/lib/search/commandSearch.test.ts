import { describe, it, expect } from "vitest";

import { COMMANDS } from "@/lib/commands/registry";
import { getEvent, speechOrder } from "@/lib/format/events";
import { searchCommands, searchSpeechCommands } from "@/lib/search/commandSearch";

const speeches = speechOrder(getEvent("policy"), "aff");
const parli = speechOrder(getEvent("parli"), "aff");
const pf = speechOrder(getEvent("pf"), "aff");

describe("searchSpeechCommands", () => {
    it("lists every speech in speaking order for an empty query", () => {
        const hits = searchSpeechCommands("", speeches);
        expect(hits.map((h) => h.speechId)).toEqual(speeches.map((s) => s.id));
        expect(hits[0].label).toBe(`Go to speech: ${speeches[0].name}`);
    });

    it("matches on the speech name", () => {
        const hits = searchSpeechCommands("2ac", speeches);
        expect(hits).toHaveLength(1);
        expect(hits[0].speechId).toBe("2ac");
    });

    it("matches on the shared prefix so the palette lists the jumps", () => {
        expect(searchSpeechCommands("go to speech", speeches)).toHaveLength(speeches.length);
    });

    it("returns nothing when a token is absent", () => {
        expect(searchSpeechCommands("9nr", speeches)).toEqual([]);
    });

    it("matches the abbreviation a column header shows", () => {
        expect(searchSpeechCommands("mgc", parli).map((h) => h.speechId)).toEqual(["mgc"]);
        expect(searchSpeechCommands("pmr", parli).map((h) => h.speechId)).toEqual(["pmr"]);
    });

    it("ranks the speech an abbreviation names above one that merely spells it", () => {
        // "ns" also hides inside "Co-ns-tructive", so the Aff and Neg
        // Constructives match too - below the Neg Summary, which is what a
        // debater typing "ns" is asking for.
        expect(searchSpeechCommands("ns", pf).map((h) => h.speechId)).toEqual(["ns", "ac", "nc"]);
    });

    it("finds a Block by either speech folded into it", () => {
        for (const q of ["2nc", "1nr"]) {
            expect(searchSpeechCommands(q, speeches).map((h) => h.speechId)).toEqual(["block"]);
        }
        for (const q of ["moc", "lor", "leader of opposition rebuttal"]) {
            expect(searchSpeechCommands(q, parli).map((h) => h.speechId)).toEqual(["block"]);
        }
    });

    it("finds a parli speech by the Policy-style name debaters also use", () => {
        expect(searchSpeechCommands("1ac", parli).map((h) => h.speechId)).toEqual(["pm"]);
        expect(searchSpeechCommands("1nc", parli).map((h) => h.speechId)).toEqual(["loc"]);
    });

    it("finds the Prime Minister by the Constructive name some circuits give it", () => {
        for (const q of ["pmc", "prime minister constructive"]) {
            expect(searchSpeechCommands(q, parli).map((h) => h.speechId)).toEqual(["pm"]);
        }
    });

    it("returns both speeches a shared role name reaches, in speaking order", () => {
        // "Leader of the Opposition Constructive" is the LOC's own name; the
        // Block carries the phrase only inside its LOR alias.
        expect(searchSpeechCommands("leader of opposition", parli).map((h) => h.speechId)).toEqual([
            "loc",
            "block",
        ]);
    });

    it("never shows an alias, only the name the speech goes by", () => {
        expect(searchSpeechCommands("mgc", parli)[0].label).toBe(
            "Go to speech: Member of the Government Constructive",
        );
        expect(searchSpeechCommands("pmc", parli)[0].label).toBe("Go to speech: Prime Minister");
    });
});

describe("searchCommands", () => {
    it("finds a command by its mark without showing the alias", () => {
        const hits = searchCommands("strikethrough");
        expect(hits.map((h) => h.id)).toEqual(["format.toggleKicked"]);
        // Keywords rank the hit, they never reach the palette row.
        expect(hits[0].label).toBe("Toggle kicked");
    });

    it("ranks a keyword hit below every command the query names", () => {
        // "Toggle bold" is a label match; kicked matches only on its keywords.
        const hits = searchCommands("bold");
        expect(hits[0].id).toBe("format.toggleBold");
    });

    it("matches on the label as before", () => {
        expect(searchCommands("toggle kicked")[0].label).toBe(
            COMMANDS["format.toggleKicked"].label,
        );
    });
});
