import { beforeEach, describe, expect, it } from "vitest";

import { seedDoc } from "@/lib/collab/doc";
import { merge } from "@/lib/collab/merge";
import { applyOp } from "@/lib/collab/ops";
import {
    adoptReplicaActor,
    clearReplica,
    getReplica,
    recordOp,
    replicaActor,
    seedReplica,
} from "@/lib/collab/replica";
import { createClock } from "@/lib/collab/stamp";
import { makeFlowRound, type FlowRound } from "@/lib/model/flow";

function round(): FlowRound {
    const r = makeFlowRound({});
    const sheet = r.sheets.find((s) => s.kind !== "cx")!;
    sheet.data = [["a"], ["b"]];
    return r;
}

beforeEach(() => {
    clearReplica();
});

describe("the actor a replica writes under", () => {
    it("is empty while solo, so two peers opening one file agree exactly", () => {
        seedReplica(round());
        expect(replicaActor()).toBe("");
    });

    it("becomes this peer's own once a session gives it an identity", () => {
        seedReplica(round());
        adoptReplicaActor("alex");
        expect(replicaActor()).toBe("alex");
    });

    it("keeps the document it already had when it adopts one", () => {
        const r = round();
        seedReplica(r);
        const sheetId = r.sheets.find((s) => s.kind !== "cx")!.id;
        recordOp({ kind: "cellText", sheetId, col: 0, row: 0, text: "typed before the session" });
        adoptReplicaActor("alex");

        const cells = Object.values(getReplica()!.sheets[sheetId].cells);
        expect(cells.map((c) => c.text)).toContain("typed before the session");
    });

    it("stamps cells written after adoption with that identity", () => {
        const r = round();
        seedReplica(r);
        adoptReplicaActor("alex");
        const sheetId = r.sheets.find((s) => s.kind !== "cx")!.id;
        recordOp({ kind: "insertCell", sheetId, col: 0, row: 1 });

        const mine = Object.values(getReplica()!.sheets[sheetId].cells).filter(
            (c) => c.actor === "alex",
        );
        expect(mine).toHaveLength(1);
    });
});

describe("two peers inserting at one position at the same moment", () => {
    it("both survive, because their cells carry different identities", () => {
        const r = round();
        const sheet = r.sheets.find((s) => s.kind !== "cx")!;
        const base = seedDoc(r);

        let t1 = 1_000;
        let t2 = 1_000;
        const alex = applyOp(
            base,
            { kind: "insertCell", sheetId: sheet.id, col: 0, row: 1 },
            { actor: "alex", clock: createClock("alex", () => t1++) },
        );
        const sam = applyOp(
            base,
            { kind: "insertCell", sheetId: sheet.id, col: 0, row: 1 },
            { actor: "sam", clock: createClock("sam", () => t2++) },
        );

        const merged = merge(alex, sam).doc;
        const alive = Object.values(merged.sheets[sheet.id].cells).filter(
            (c) => c.deleted === null,
        );
        // Two seeded cells plus one insert from each peer. With one shared
        // identity the two inserts collide on the same key and one is lost
        // with nothing reporting it.
        expect(alive).toHaveLength(4);
    });
});
