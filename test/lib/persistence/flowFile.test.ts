import { describe, expect, it } from "vitest";

import { makeFlowRound, type FlowRound } from "@/lib/model/flow";
import {
    FLOW_FILE_VERSION,
    MAX_FLOW_TEXT_CHARS,
    MAX_ROUND_CELLS,
    parseFlowFile,
    parseLegacyExport,
    serializeFlow,
} from "@/lib/persistence/flowFile";

function fileFor(round: FlowRound): string {
    return serializeFlow(round);
}

/** A valid envelope with one field of the round replaced. */
function withRound(patch: Record<string, unknown>): string {
    const round = { ...makeFlowRound({}), ...patch };
    return JSON.stringify({ version: FLOW_FILE_VERSION, round });
}

describe("serializeFlow", () => {
    it("writes a versioned envelope as readable JSON", () => {
        const round = makeFlowRound({ event: "ld" });
        const parsed: unknown = JSON.parse(serializeFlow(round));
        expect(parsed).toMatchObject({ version: FLOW_FILE_VERSION });
        expect(serializeFlow(round)).toContain("\n  ");
        expect(serializeFlow(round).endsWith("\n")).toBe(true);
    });

    it("refuses a round this parser could not read back", () => {
        // A round projected from a replica carries whatever a peer put in a
        // register, and a file written from one of those is a round the debater
        // can never open again. The write is the last place to refuse.
        const round = makeFlowRound({});
        const unknownEvent = { ...round, event: "moot-court" as FlowRound["event"] };
        const noSide = { ...round, firstSide: "nobody" as FlowRound["firstSide"] };
        expect(() => serializeFlow(unknownEvent)).toThrow(/event is not a known debate event/);
        expect(() => serializeFlow(noSide)).toThrow(/round\.firstSide is not "aff" or "neg"/);
    });

    it("refuses an event name that only the prototype chain answers to", () => {
        // `in` walks the chain, so "constructor" would pass the check and then
        // name no event at all wherever the table is indexed.
        const round = { ...makeFlowRound({}), event: "constructor" as FlowRound["event"] };
        expect(() => serializeFlow(round)).toThrow(/event is not a known debate event/);
        expect(() => parseFlowFile(withRound({ event: "constructor" }))).toThrow(
            /round\.event is not a known debate event/,
        );
    });

    /**
     * The shell refuses a flow by the bytes on disk, and one character is up to
     * three of them, so a round of Chinese well inside the character cap is a
     * file the app writes and then cannot read back - the round gone, which is
     * the whole reason the write refuses ahead of the read. Only a bug in the
     * projection's own byte budget gets here, which is what a last line of
     * defence is for.
     */
    it("refuses a round whose characters fit the file and whose bytes do not", () => {
        const round = makeFlowRound({});
        const flow = round.sheets.find((s) => s.kind !== "cx")!;
        // Three bytes a character, so two thirds of the cap in characters is
        // twice the cap in bytes.
        flow.data = [["中".repeat(Math.floor((2 * MAX_FLOW_TEXT_CHARS) / 3))]];
        expect(JSON.stringify(round).length).toBeLessThan(MAX_FLOW_TEXT_CHARS);
        expect(() => serializeFlow(round)).toThrow(/longer than/);
    }, 30_000);
});

describe("parseFlowFile", () => {
    it("preserves identity, because opening a file is not importing one", () => {
        const round = makeFlowRound({ event: "pf", firstSide: "neg" });
        const reopened = parseFlowFile(fileFor(round));

        expect(reopened.id).toBe(round.id);
        expect(reopened.createdAt).toBe(round.createdAt);
        expect(reopened.updatedAt).toBe(round.updatedAt);
        expect(reopened.event).toBe("pf");
        expect(reopened.firstSide).toBe("neg");
    });

    it("round-trips sheet content and cell metadata", () => {
        const round = makeFlowRound({});
        round.sheets[1].data = [
            ["extinction", null],
            [null, "turn"],
        ];
        round.sheets[1].meta = { "0,0": { bold: true, card: true } };

        const reopened = parseFlowFile(fileFor(round));
        expect(reopened.sheets[1].data).toEqual([
            ["extinction", null],
            [null, "turn"],
        ]);
        expect(reopened.sheets[1].meta["0,0"]).toEqual({ bold: true, card: true });
    });

    it("drops the legacy soft-delete flag rather than carrying it forward", () => {
        const text = withRound({ deletedAt: 1234 });
        expect(parseFlowFile(text)).not.toHaveProperty("deletedAt");
    });

    it("tells the user to update when the file is from a newer ebb", () => {
        const text = JSON.stringify({ version: FLOW_FILE_VERSION + 1, round: makeFlowRound({}) });
        expect(() => parseFlowFile(text)).toThrow(/newer version of ebb/);
    });

    it("refuses the retired node-model versions", () => {
        const text = JSON.stringify({ version: 2, round: makeFlowRound({}) });
        expect(() => parseFlowFile(text)).toThrow(/retired format/);
    });

    it("refuses a backup rather than opening its first round by accident", () => {
        const text = JSON.stringify({
            version: FLOW_FILE_VERSION,
            kind: "backup",
            rounds: [makeFlowRound({})],
        });
        expect(() => parseFlowFile(text)).toThrow(/backup/);
    });

    it("names what is wrong, so a corrupt file is diagnosable", () => {
        expect(() => parseFlowFile("{oh no")).toThrow(/not valid JSON/);
        expect(() => parseFlowFile(withRound({ id: 7 }))).toThrow(
            "Invalid flow file: round.id is not a string",
        );
        expect(() => parseFlowFile(withRound({ event: "moot-court" }))).toThrow(
            /round\.event is not a known debate event/,
        );
        expect(() => parseFlowFile(withRound({ scouting: {} }))).toThrow(
            /round\.scouting\.aff is not an object/,
        );
    });

    it("rejects a truncated grid at the cell that is wrong", () => {
        const round = makeFlowRound({});
        const broken = JSON.parse(serializeFlow(round)) as { round: { sheets: unknown[] } };
        (broken.round.sheets[1] as { data: unknown }).data = [["ok"], [3]];

        expect(() => parseFlowFile(JSON.stringify(broken))).toThrow(
            /sheets\[1\]\.data\[1\]\[0\] is not text or null/,
        );
    });

    it("refuses a grid whose claimed cells no machine can materialize", () => {
        // Rows and width are independent dimensions, both taken from the file,
        // and every consumer allocates their product: one wide row beside many
        // single-cell rows is 10^10 cells from under a megabyte on disk.
        const raw = JSON.parse(serializeFlow(makeFlowRound({}))) as {
            round: { sheets: { data: unknown }[] };
        };
        raw.round.sheets[1].data = [
            Array.from({ length: 20_000 }, () => ""),
            ...Array.from({ length: 200 }, () => [""]),
        ];

        expect(() => parseFlowFile(JSON.stringify(raw))).toThrow(
            new RegExp(`sheets hold more than ${MAX_ROUND_CELLS} cells`),
        );
    });

    it("refuses text too large to be a flow before it parses it", () => {
        expect(() => parseFlowFile("x".repeat(MAX_FLOW_TEXT_CHARS + 1))).toThrow(
            /too large to be one/,
        );
    });

    it("refuses a cell-meta key that is not a cell, since it reaches the grid as one", () => {
        const raw = JSON.parse(serializeFlow(makeFlowRound({}))) as {
            round: { sheets: { meta: unknown }[] };
        };
        // Parsed rather than written as a literal: `__proto__` in an object
        // literal is the prototype setter, and the file gives an own key.
        raw.round.sheets[1].meta = JSON.parse('{"__proto__":{"bold":true}}');
        expect(() => parseFlowFile(JSON.stringify(raw))).toThrow(
            /meta\["__proto__"\] is not a "row,col" cell/,
        );

        raw.round.sheets[1].meta = { banana: { bold: true } };
        expect(() => parseFlowFile(JSON.stringify(raw))).toThrow(
            /meta\["banana"\] is not a "row,col" cell/,
        );
    });

    it("accepts a sheet that predates cell metadata", () => {
        const round = makeFlowRound({});
        const raw = JSON.parse(serializeFlow(round)) as { round: { sheets: unknown[] } };
        delete (raw.round.sheets[0] as { meta?: unknown }).meta;

        expect(parseFlowFile(JSON.stringify(raw)).sheets[0].meta).toEqual({});
    });
});

describe("parseLegacyExport", () => {
    it("mints a fresh identity, because an export was a snapshot not a document", () => {
        const round = makeFlowRound({});
        const [imported] = parseLegacyExport(JSON.stringify({ version: FLOW_FILE_VERSION, round }));
        expect(imported.id).not.toBe(round.id);
    });

    it("explodes a multi-round backup into separate rounds", () => {
        const a = makeFlowRound({ event: "policy" });
        const b = makeFlowRound({ event: "ld" });
        const rounds = parseLegacyExport(
            JSON.stringify({ version: FLOW_FILE_VERSION, kind: "backup", rounds: [a, b] }),
        );

        expect(rounds).toHaveLength(2);
        expect(rounds.map((r) => r.event)).toEqual(["policy", "ld"]);
        expect(new Set(rounds.map((r) => r.id)).size).toBe(2);
    });

    it("points at the offending round inside a backup", () => {
        const text = JSON.stringify({
            version: FLOW_FILE_VERSION,
            kind: "backup",
            rounds: [makeFlowRound({}), { id: "x", sheets: [] }],
        });
        expect(() => parseLegacyExport(text)).toThrow(/rounds\[1\]/);
    });
});
