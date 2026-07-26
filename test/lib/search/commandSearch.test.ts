import { describe, it, expect } from "vitest";

import { getEvent, speechOrder } from "@/lib/format/events";
import { searchSpeechCommands } from "@/lib/search/commandSearch";

const speeches = speechOrder(getEvent("policy"), "aff");

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
});
