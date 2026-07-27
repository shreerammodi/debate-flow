import { beforeEach, describe, expect, it } from "vitest";

import { projectSheet, seedDoc } from "@/lib/collab/doc";
import { merge } from "@/lib/collab/merge";
import { applyOp, type CollabOp, type OpContext } from "@/lib/collab/ops";
import {
    driftedSheetIds,
    getReplica,
    healReplica,
    recordOp,
    replaceReplicaDoc,
    seedReplica,
    setLocalChangeListener,
} from "@/lib/collab/replica";
import { createClock } from "@/lib/collab/stamp";
import type { CollabDoc } from "@/lib/collab/types";
import { trimGrid } from "@/lib/grid/codec";
import { makeFlowRound, type FlowRound } from "@/lib/model/flow";

/** How wide a policy flow sheet's grid is, which is what the store stores. */
const GRID_COLS = 8;

let round: FlowRound;
let sheetId: string;
/** The replica both machines derive from the file, before either types. */
let pristine: CollabDoc;

beforeEach(() => {
    setLocalChangeListener(null);
    round = makeFlowRound({});
    const sheet = round.sheets.find((s) => s.kind !== "cx")!;
    sheet.data = [];
    sheetId = sheet.id;
    pristine = seedDoc(round);
    seedReplica(round, "alex");
});

/**
 * The store copy HotGrid's snapshot writes: every grid column present, only
 * trailing empty rows cut.
 */
function snapshot(cells: Record<string, string>): void {
    let height = 0;
    for (const key of Object.keys(cells)) height = Math.max(height, Number(key.split(",")[0]) + 1);
    const data: (string | null)[][] = [];
    for (let r = 0; r < Math.max(height, 250); r++) {
        const line: (string | null)[] = [];
        for (let c = 0; c < GRID_COLS; c++) line.push(cells[`${r},${c}`] ?? null);
        data.push(line);
    }
    round.sheets.find((s) => s.id === sheetId)!.data = trimGrid(data);
}

function rows(doc: CollabDoc, col: number): (string | null)[] {
    return projectSheet(doc.sheets[sheetId]).data.map((r) => r[col] ?? null);
}

describe("a sheet the debater is typing into", () => {
    it("reports no drift, so the autosave never re-seeds it", () => {
        recordOp({ kind: "cellText", sheetId, col: 0, row: 6, text: "arg" });
        snapshot({ "6,0": "arg" });
        expect(driftedSheetIds(round)).toEqual([]);
    });

    it("reports no drift once a second column is typed into", () => {
        recordOp({ kind: "cellText", sheetId, col: 0, row: 6, text: "arg" });
        recordOp({ kind: "cellText", sheetId, col: 1, row: 8, text: "blah" });
        snapshot({ "6,0": "arg", "8,1": "blah" });
        expect(driftedSheetIds(round)).toEqual([]);
    });
});

describe("the repair that re-keys a sheet", () => {
    it("runs when nobody is sharing the round", () => {
        recordOp({ kind: "cellText", sheetId, col: 0, row: 0, text: "arg" });
        // A store the replica never heard about: the drift the repair is for.
        snapshot({ "0,0": "arg", "1,0": "typed behind the replica's back" });
        expect(healReplica(round)).toEqual([sheetId]);
        expect(rows(getReplica()!, 0)).toEqual(["arg", "typed behind the replica's back"]);
    });

    // Re-keying every cell from its row position is invisible on one machine
    // and unrecoverable on two: the peer still holds the old keys, and the
    // merge keeps both sets.
    it("never runs while a session is live", () => {
        recordOp({ kind: "cellText", sheetId, col: 0, row: 0, text: "arg" });
        const before = getReplica();
        snapshot({ "0,0": "arg", "1,0": "typed behind the replica's back" });

        setLocalChangeListener(() => {});
        expect(healReplica(round)).toEqual([]);
        expect(getReplica()).toBe(before);
    });
});

describe("two machines flowing the same round", () => {
    /**
     * The partner's replica of the same file, with writes of their own. Seeded
     * from the round as both machines opened it, which is what a join hands
     * the guest and what makes the first merge a merge.
     */
    function partner(...ops: CollabOp[]): CollabDoc {
        let t = 5_000;
        const ctx: OpContext = { actor: "sam", clock: createClock("sam", () => t++) };
        let doc = pristine;
        for (const op of ops) doc = applyOp(doc, op, ctx);
        return doc;
    }

    it("hold the same sheet after a save lands between two edits", () => {
        setLocalChangeListener(() => {});
        recordOp({ kind: "cellText", sheetId, col: 0, row: 0, text: "arg" });
        recordOp({ kind: "cellText", sheetId, col: 0, row: 1, text: "garg" });
        snapshot({ "0,0": "arg", "1,0": "garg" });

        // The autosave fires mid-round, between this machine's writes and the
        // partner's arriving.
        healReplica(round);

        const theirs = partner({ kind: "cellText", sheetId, col: 1, row: 1, text: "blah" });
        const mineAfter = merge(getReplica()!, theirs).doc;
        const theirsAfter = merge(theirs, getReplica()!).doc;
        replaceReplicaDoc(mineAfter);

        expect(rows(mineAfter, 0)).toEqual(rows(theirsAfter, 0));
        expect(rows(mineAfter, 1)).toEqual(rows(theirsAfter, 1));
        expect(rows(mineAfter, 0)).toEqual(["arg", "garg"]);
        expect(rows(mineAfter, 1)).toEqual([null, "blah"]);
    });

    it("agree cell for cell however many saves and exchanges interleave", () => {
        setLocalChangeListener(() => {});
        let theirs = pristine;
        let t = 5_000;
        const sam: OpContext = { actor: "sam", clock: createClock("sam", () => t++) };

        for (let i = 0; i < 6; i++) {
            recordOp({ kind: "cellText", sheetId, col: 0, row: i, text: `arg ${i}` });
            snapshot(Object.fromEntries([...Array(i + 1)].map((_, r) => [`${r},0`, `arg ${r}`])));
            healReplica(round);

            theirs = applyOp(
                theirs,
                { kind: "cellText", sheetId, col: 1, row: i, text: `blah ${i}` },
                sam,
            );
            replaceReplicaDoc(merge(getReplica()!, theirs).doc);
            theirs = merge(theirs, getReplica()!).doc;
        }

        expect(projectSheet(getReplica()!.sheets[sheetId]).data).toEqual(
            projectSheet(theirs.sheets[sheetId]).data,
        );
        expect(rows(getReplica()!, 0)).toEqual([
            "arg 0",
            "arg 1",
            "arg 2",
            "arg 3",
            "arg 4",
            "arg 5",
        ]);
        expect(rows(getReplica()!, 1)).toEqual([
            "blah 0",
            "blah 1",
            "blah 2",
            "blah 3",
            "blah 4",
            "blah 5",
        ]);
    });
});
