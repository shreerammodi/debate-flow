import { beforeEach, describe, expect, it } from "vitest";

import { seedDoc } from "@/lib/collab/doc";
import { isForeignSource, sourceOwner } from "@/lib/collab/foreignSource";
import { applyOp, type OpContext } from "@/lib/collab/ops";
import { createClock } from "@/lib/collab/stamp";
import type { CollabDoc } from "@/lib/collab/types";
import { makeFlowRound, type FlowRound } from "@/lib/model/flow";

const ME = "me-endpoint";
const THEM = "them-endpoint";

const SOURCE = { app: "cardmirror", token: "tok-1", key: "k", title: "Aff Core" };

let round: FlowRound;
let sheetId: string;
let doc: CollabDoc;

function ctx(actor: string, start: number): OpContext {
    let t = start;
    return { actor, clock: createClock(actor, () => t++) };
}

beforeEach(() => {
    round = makeFlowRound({});
    const flow = round.sheets.find((s) => s.kind !== "cx")!;
    sheetId = flow.id;
    flow.data = [["a0", "b0"]];
    doc = seedDoc(round);
});

describe("sourceOwner", () => {
    it("names the peer whose meta carried the source", () => {
        const next = applyOp(
            doc,
            { kind: "cellMeta", sheetId, col: 0, row: 0, meta: { source: SOURCE } },
            ctx(THEM, 9_000),
        );
        expect(sourceOwner(next.sheets[sheetId], 0, 0)).toBe(THEM);
    });

    it("names nobody for a cell with no source at all", () => {
        const next = applyOp(
            doc,
            { kind: "cellMeta", sheetId, col: 0, row: 0, meta: { bold: true } },
            ctx(THEM, 9_000),
        );
        expect(sourceOwner(next.sheets[sheetId], 0, 0)).toBeNull();
    });

    it("names nobody for a source that came in from the file", () => {
        // Seeded meta carries the origin stamp, which belongs to no peer.
        expect(sourceOwner(doc.sheets[sheetId], 0, 0)).toBeNull();
    });

    it("names nobody for a cell that is not there", () => {
        expect(sourceOwner(doc.sheets[sheetId], 9, 9)).toBeNull();
        expect(sourceOwner(undefined, 0, 0)).toBeNull();
    });
});

describe("isForeignSource", () => {
    it("calls a partner's token foreign, because it means nothing here", () => {
        expect(isForeignSource(THEM, ME)).toBe(true);
    });

    it("calls my own token mine", () => {
        expect(isForeignSource(ME, ME)).toBe(false);
    });

    it("calls an unowned token mine, which is every solo round", () => {
        expect(isForeignSource(null, ME)).toBe(false);
        expect(isForeignSource("", ME)).toBe(false);
    });
});
