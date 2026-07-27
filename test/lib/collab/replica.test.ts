import { beforeEach, describe, expect, it } from "vitest";

import { projectDoc, seedDoc } from "@/lib/collab/doc";
import { applyOp, type OpContext } from "@/lib/collab/ops";
import {
    clearReplica,
    driftedSheetIds,
    getReplica,
    healReplica,
    recordOp,
    replaceReplicaDoc,
    replicaActor,
    replicaRoundId,
    resyncSheet,
    seedReplica,
    setLocalChangeListener,
} from "@/lib/collab/replica";
import { compareStamps, createClock } from "@/lib/collab/stamp";
import { makeFlowRound, type FlowRound } from "@/lib/model/flow";

function round(): FlowRound {
    const r = makeFlowRound({});
    const flow = r.sheets.find((s) => s.kind !== "cx")!;
    flow.data = [
        ["perm", "link"],
        ["cap bad", "turn"],
    ];
    return r;
}

beforeEach(() => {
    clearReplica();
});

describe("seedReplica", () => {
    it("holds nothing until a round opens", () => {
        expect(getReplica()).toBeNull();
        expect(replicaRoundId()).toBeNull();
    });

    it("seeds deterministically from the round", () => {
        const r = round();
        seedReplica(r);
        expect(getReplica()).toEqual(seedDoc(r));
        expect(replicaRoundId()).toBe(r.id);
    });

    it("replaces the previous round outright, with no close in between", () => {
        const first = round();
        seedReplica(first);
        const second = round();
        seedReplica(second);
        expect(replicaRoundId()).toBe(second.id);
        expect(getReplica()).toEqual(seedDoc(second));
    });

    it("adopts a doc recovered from a sidecar instead of seeding", () => {
        const r = round();
        const recovered = seedDoc(r);
        recovered.round.event = { value: "pf", stamp: { ms: 9, counter: 0, actor: "sam" } };
        seedReplica(r, "alex", recovered);
        expect(getReplica()!.round.event.value).toBe("pf");
    });

    it("clears back to nothing", () => {
        seedReplica(round());
        clearReplica();
        expect(getReplica()).toBeNull();
    });
});

describe("recordOp", () => {
    it("applies an op to the live replica", () => {
        const r = round();
        const sheetId = r.sheets.find((s) => s.kind !== "cx")!.id;
        seedReplica(r);
        recordOp({ kind: "cellText", sheetId, col: 0, row: 0, text: "perm, then CP" });
        const data = projectDoc(getReplica()!, r).sheets.find((s) => s.id === sheetId)!.data;
        expect(data[0][0]).toBe("perm, then CP");
    });

    it("is a no-op with no round open, so a stray hook cannot throw", () => {
        expect(() =>
            recordOp({ kind: "cellText", sheetId: "x", col: 0, row: 0, text: "a" }),
        ).not.toThrow();
    });

    it("stamps two writes in order", () => {
        const r = round();
        const sheetId = r.sheets.find((s) => s.kind !== "cx")!.id;
        seedReplica(r);
        recordOp({ kind: "cellText", sheetId, col: 0, row: 0, text: "one" });
        const first = Object.values(getReplica()!.sheets[sheetId].cells).find(
            (c) => c.text === "one",
        )!.textStamp;
        recordOp({ kind: "cellText", sheetId, col: 0, row: 0, text: "two" });
        const second = Object.values(getReplica()!.sheets[sheetId].cells).find(
            (c) => c.text === "two",
        )!.textStamp;
        expect(second.ms > first.ms || second.counter > first.counter).toBe(true);
    });
});

describe("resyncSheet", () => {
    it("re-seeds one sheet from the store copy and leaves the others alone", () => {
        const r = round();
        const flow = r.sheets.find((s) => s.kind !== "cx")!;
        const cx = r.sheets.find((s) => s.kind === "cx")!;
        seedReplica(r);
        const cxBefore = getReplica()!.sheets[cx.id];

        resyncSheet({
            ...flow,
            data: [
                ["moved", "link"],
                ["cap bad", "turn"],
            ],
        });
        const data = projectDoc(getReplica()!, r).sheets.find((s) => s.id === flow.id)!.data;
        expect(data[0][0]).toBe("moved");
        expect(getReplica()!.sheets[cx.id]).toBe(cxBefore);
    });

    it("is a no-op with no round open", () => {
        expect(() => resyncSheet(round().sheets[0])).not.toThrow();
    });
});

describe("self-heal", () => {
    it("reports no drift on a replica that tracked every edit", () => {
        const r = round();
        seedReplica(r);
        expect(driftedSheetIds(r)).toEqual([]);
    });

    it("reports no drift when the store row is merely ragged", () => {
        const r = round();
        seedReplica(r);
        const flow = r.sheets.find((s) => s.kind !== "cx")!;
        flow.data = [["perm", "link"], ["cap bad", "turn"], []];
        expect(driftedSheetIds(r)).toEqual([]);
    });

    it("spots a sheet the replica missed, and repairs it", () => {
        const r = round();
        const flow = r.sheets.find((s) => s.kind !== "cx")!;
        seedReplica(r);
        // A hook that never fired: the store moved on, the replica did not.
        flow.data = [
            ["perm", "link"],
            ["cap bad", "MISSED"],
        ];
        expect(driftedSheetIds(r)).toEqual([flow.id]);
        expect(healReplica(r)).toEqual([flow.id]);
        expect(driftedSheetIds(r)).toEqual([]);
        const data = projectDoc(getReplica()!, r).sheets.find((s) => s.id === flow.id)!.data;
        expect(data[1][1]).toBe("MISSED");
    });

    it("adopts a sheet the replica has never seen", () => {
        const r = round();
        seedReplica(r);
        const added = { ...r.sheets[0], id: "sheet-new", title: "DA", data: [["shell"]], meta: {} };
        r.sheets = [...r.sheets, added];
        expect(driftedSheetIds(r)).toEqual(["sheet-new"]);
        healReplica(r);
        expect(getReplica()!.sheets["sheet-new"]).toBeDefined();
    });

    it("heals nothing and reports nothing with no round open", () => {
        expect(driftedSheetIds(round())).toEqual([]);
        expect(healReplica(round())).toEqual([]);
    });
});

describe("the local-change bridge", () => {
    it("tells the listener about every write a session has to push", () => {
        const r = round();
        seedReplica(r);
        let told = 0;
        setLocalChangeListener(() => told++);

        recordOp({ kind: "cellText", sheetId: r.sheets[0].id, col: 0, row: 0, text: "typed" });
        expect(told).toBe(1);
        resyncSheet(r.sheets[0]);
        expect(told).toBe(2);

        setLocalChangeListener(null);
        recordOp({ kind: "cellText", sheetId: r.sheets[0].id, col: 0, row: 0, text: "after" });
        expect(told).toBe(2);
    });

    it("stays quiet for a write with no round open", () => {
        let told = 0;
        setLocalChangeListener(() => told++);
        recordOp({ kind: "cellText", sheetId: "gone", col: 0, row: 0, text: "typed" });
        setLocalChangeListener(null);
        expect(told).toBe(0);
    });
});

describe("replaceReplicaDoc", () => {
    it("keeps this machine's identity and clock across a merge", () => {
        const r = round();
        seedReplica(r, "alex");
        const sheetId = r.sheets.find((s) => s.kind !== "cx")!.id;
        recordOp({ kind: "cellText", sheetId, col: 0, row: 0, text: "mine" });
        const first = getReplica()!.sheets[sheetId].cells;
        const stampBefore = Object.values(first).find((c) => c.text === "mine")!.textStamp;

        replaceReplicaDoc(getReplica()!);
        expect(replicaActor()).toBe("alex");

        recordOp({ kind: "cellText", sheetId, col: 0, row: 0, text: "mine again" });
        const stampAfter = Object.values(getReplica()!.sheets[sheetId].cells).find(
            (c) => c.text === "mine again",
        )!.textStamp;
        expect(compareStamps(stampAfter, stampBefore)).toBeGreaterThan(0);
    });

    it("raises the clock past a peer sitting in the future", () => {
        const r = round();
        seedReplica(r, "alex");
        const sheetId = r.sheets.find((s) => s.kind !== "cx")!.id;
        // Pinned, not read again at the end: the wall clock moving between the
        // peer's write and the assertion would decide whether this passes.
        const future = Date.now() + 60_000;
        const ahead: OpContext = { actor: "sam", clock: createClock("sam", () => future) };
        replaceReplicaDoc(
            applyOp(getReplica()!, { kind: "cellText", sheetId, col: 0, row: 0, text: "" }, ahead),
        );

        recordOp({ kind: "cellText", sheetId, col: 1, row: 0, text: "mine" });
        const mine = Object.values(getReplica()!.sheets[sheetId].cells).find(
            (c) => c.text === "mine",
        )!;
        expect(mine.textStamp.ms).toBe(future);
    });
});
