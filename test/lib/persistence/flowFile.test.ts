import { describe, expect, it } from "vitest";

import { makeFlowRound, type FlowRound } from "@/lib/model/flow";
import {
    FLOW_FILE_VERSION,
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
        expect(() => parseFlowFile(withRound({ event: "parli" }))).toThrow(
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
