import { describe, expect, it } from "vitest";

import { deltaSince, emptyVector } from "@/lib/collab/delta";
import { seedDoc } from "@/lib/collab/doc";
import { applyOp, type CollabOp, type OpContext } from "@/lib/collab/ops";
import { createShadow, type Shadow } from "@/lib/collab/shadow";
import { createClock } from "@/lib/collab/stamp";
import type { CollabDoc } from "@/lib/collab/types";
import { makeFlowRound, type FlowRound } from "@/lib/model/flow";

function peer(actor: string, start: number): OpContext {
    let t = start;
    return { actor, clock: createClock(actor, () => t++) };
}

function base(): { round: FlowRound; doc: CollabDoc; sheetId: string } {
    const round = makeFlowRound({});
    const flow = round.sheets.find((s) => s.kind !== "cx")!;
    flow.data = [
        ["perm", "link"],
        ["cap bad", "turn"],
    ];
    return { round, doc: seedDoc(round), sheetId: flow.id };
}

/**
 * The host's live replica behind a shadow reading it fresh, plus a partner
 * writing against the same seed the way a real guest does.
 */
function harness(): {
    round: FlowRound;
    sheetId: string;
    shadow: Shadow;
    live: () => CollabDoc;
    host: (op: CollabOp) => void;
    partner: (...ops: CollabOp[]) => CollabDoc;
} {
    const { round, doc, sheetId } = base();
    let live = doc;
    let tick = 0;
    const sam = peer("sam", 500);
    return {
        round,
        sheetId,
        shadow: createShadow({ doc: () => live, base: () => round, now: () => ++tick }),
        live: () => live,
        host: (op) => {
            live = applyOp(live, op, peer("host", 900));
        },
        partner: (...ops) => ops.reduce((d, op) => applyOp(d, op, sam), doc),
    };
}

const write = (sheetId: string, col: number, row: number, text: string): CollabOp => ({
    kind: "cellText",
    sheetId,
    col,
    row,
    text,
});

describe("createShadow", () => {
    it("reports a cell the partner would have changed, with both texts", () => {
        const { shadow, sheetId, partner } = harness();
        const entry = shadow.observe("sam", partner(write(sheetId, 1, 0, "no link")));
        expect(entry.diffs).toEqual([{ sheetId, col: 1, row: 0, mine: "link", theirs: "no link" }]);
        expect(entry.from).toBe("sam");
    });

    it("reports nothing when the remote change is one this machine already has", () => {
        const { shadow, sheetId, live, host } = harness();
        host(write(sheetId, 1, 0, "no link"));
        const entry = shadow.observe("sam", live());
        expect(entry.diffs).toEqual([]);
        expect(entry.dropped).toEqual([]);
    });

    it("reports a cell the partner would have added where this machine has nothing", () => {
        const { shadow, sheetId, partner } = harness();
        const entry = shadow.observe("sam", partner(write(sheetId, 0, 2, "new arg")));
        expect(entry.diffs).toEqual([{ sheetId, col: 0, row: 2, mine: "", theirs: "new arg" }]);
    });

    it("carries the merge's dropped cells onto the entry", () => {
        const { shadow, sheetId, host, partner } = harness();
        host(write(sheetId, 0, 0, "keep me"));
        const removed = partner({ kind: "removeRow", sheetId, row: 0 });
        const entry = shadow.observe("sam", removed);
        expect(entry.dropped.map((d) => ({ text: d.text, col: d.col, by: d.deletedBy }))).toEqual([
            { text: "keep me", col: 0, by: "sam" },
            { text: "link", col: 1, by: "sam" },
        ]);
        expect(entry.diffs).toEqual([
            { sheetId, col: 0, row: 0, mine: "keep me", theirs: "cap bad" },
            { sheetId, col: 1, row: 0, mine: "link", theirs: "turn" },
            { sheetId, col: 0, row: 1, mine: "cap bad", theirs: "" },
            { sheetId, col: 1, row: 1, mine: "turn", theirs: "" },
        ]);
    });

    it("reports every cell on a sheet the partner would have removed", () => {
        const { shadow, sheetId, partner } = harness();
        const entry = shadow.observe("sam", partner({ kind: "removeSheet", sheetId }));
        expect(entry.diffs).toEqual([
            { sheetId, col: 0, row: 0, mine: "perm", theirs: "" },
            { sheetId, col: 1, row: 0, mine: "link", theirs: "" },
            { sheetId, col: 0, row: 1, mine: "cap bad", theirs: "" },
            { sheetId, col: 1, row: 1, mine: "turn", theirs: "" },
        ]);
    });

    it("never mutates the live document or the incoming one", () => {
        const { shadow, sheetId, host, live, partner } = harness();
        host(write(sheetId, 0, 0, "mine"));
        const incoming = partner(write(sheetId, 0, 0, "theirs"), {
            kind: "removeRow",
            sheetId,
            row: 1,
        });
        const liveBefore = structuredClone(live());
        const incomingBefore = structuredClone(incoming);
        shadow.observe("sam", incoming);
        expect(live()).toEqual(liveBefore);
        expect(incoming).toEqual(incomingBefore);
    });

    it("reads nothing at construction, so a session can outlive an unloaded round", () => {
        const absent = () => {
            throw new Error("no round is open");
        };
        const shadow = createShadow({ doc: absent, base: absent });
        expect(shadow.entries()).toEqual([]);
    });

    it("ignores this machine's writes made before the first observation", () => {
        const { shadow, sheetId, host, partner } = harness();
        host(write(sheetId, 0, 0, "host typed this"));
        const entry = shadow.observe("sam", partner(write(sheetId, 1, 0, "no link")));
        expect(entry.diffs).toEqual([{ sheetId, col: 1, row: 0, mine: "link", theirs: "no link" }]);
    });

    it("never reports this machine's own writes when a delta does not echo them", () => {
        const { shadow, sheetId, host, live, partner } = harness();
        shadow.observe("sam", live());
        host(write(sheetId, 0, 1, "late idea"));
        // What a partner actually pushes mid-round: its own cells and no more.
        const delta = deltaSince(partner(write(sheetId, 1, 1, "turn back")), emptyVector());
        expect(Object.keys(delta.sheets[sheetId].cells)).toHaveLength(1);
        expect(shadow.observe("sam", delta).diffs).toEqual([
            { sheetId, col: 1, row: 1, mine: "turn", theirs: "turn back" },
        ]);
    });

    it("never blames the partner for a row this machine deleted itself", () => {
        const { shadow, sheetId, host, live, partner } = harness();
        shadow.observe("sam", live());
        host({ kind: "removeRow", sheetId, row: 0 });
        const delta = deltaSince(partner(write(sheetId, 1, 1, "turn back")), emptyVector());
        expect(shadow.observe("sam", delta).dropped).toEqual([]);
    });

    it("accumulates entries oldest first across several observations", () => {
        const { shadow, sheetId, partner } = harness();
        shadow.observe("sam", partner(write(sheetId, 1, 0, "no link")));
        shadow.observe("kai", partner(write(sheetId, 0, 1, "extinction")));
        expect(shadow.entries().map((e) => e.from)).toEqual(["sam", "kai"]);
        expect(shadow.entries()[0].at).toBeLessThan(shadow.entries()[1].at);
    });

    it("records an entry even when the merge changes nothing visible", () => {
        const { shadow, live } = harness();
        const entry = shadow.observe("sam", live());
        expect(shadow.entries()).toEqual([entry]);
        expect(entry.diffs).toEqual([]);
    });

    it("empties the log on clear, and reports the same diff on a repeat", () => {
        const { shadow, sheetId, partner } = harness();
        const incoming = partner(write(sheetId, 1, 0, "no link"));
        const first = shadow.observe("sam", incoming);
        shadow.clear();
        expect(shadow.entries()).toEqual([]);
        const again = shadow.observe("sam", incoming);
        expect(again.diffs).toEqual(first.diffs);
        expect(shadow.entries()).toHaveLength(1);
    });

    it("re-bases on clear, so the shadow forgets earlier remote changes", () => {
        const { shadow, sheetId, partner } = harness();
        shadow.observe("sam", partner(write(sheetId, 1, 0, "no link")));
        const second = partner(write(sheetId, 0, 1, "extinction"));
        expect(shadow.observe("sam", second).diffs).toHaveLength(2);
        shadow.clear();
        expect(shadow.observe("sam", second).diffs).toEqual([
            { sheetId, col: 0, row: 1, mine: "cap bad", theirs: "extinction" },
        ]);
    });

    it("stamps an entry with the wall clock when no clock is injected", () => {
        const { round, sheetId, partner } = harness();
        const shadow = createShadow({ doc: () => seedDoc(round), base: () => round });
        const before = Date.now();
        const entry = shadow.observe("sam", partner(write(sheetId, 1, 0, "no link")));
        expect(entry.at).toBeGreaterThanOrEqual(before);
        expect(entry.at).toBeLessThanOrEqual(Date.now());
    });
});
