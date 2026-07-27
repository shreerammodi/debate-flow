import { beforeEach, describe, expect, it, vi } from "vitest";

const warnings: string[] = [];

vi.mock("sonner", () => ({
    toast: Object.assign(() => {}, {
        warning: (m: string) => warnings.push(m),
        error: () => {},
        success: () => {},
        info: () => {},
    }),
}));

import { applyOp } from "@/lib/collab/ops";
import { getReplica, replicaActor, seedReplica } from "@/lib/collab/replica";
import { applyRemoteDoc, endSession, startForRound } from "@/lib/collab/runtime";
import { createClock } from "@/lib/collab/stamp";
import { makeFlowRound, type FlowRound } from "@/lib/model/flow";
import { useFlowStore } from "@/lib/store/useFlowStore";

function round(): FlowRound {
    const r = makeFlowRound({});
    const sheet = r.sheets.find((s) => s.kind !== "cx")!;
    sheet.data = [["perm do both"], ["turn"]];
    return r;
}

beforeEach(async () => {
    warnings.length = 0;
    await endSession();
    useFlowStore.setState({
        collabEnabled: true,
        collabRelayEnabled: true,
        shadowMode: false,
        contacts: { sam: { name: "Sam", role: "partner" } },
    });
});

describe("the live apply path", () => {
    it("puts a partner's buried write in front of the user", async () => {
        const r = round();
        const sheet = r.sheets.find((s) => s.kind !== "cx")!;
        seedReplica(r, "me");

        // The partner deletes the row this flow already had text in. Nothing
        // on the grid marks the absence, so only a message can tell the user.
        let t = 9_000;
        const theirs = applyOp(
            getReplica()!,
            { kind: "removeRow", sheetId: sheet.id, row: 0 },
            { actor: "sam", clock: createClock("sam", () => t++) },
        );

        const dropped = applyRemoteDoc(r, theirs);
        expect(dropped.length).toBeGreaterThan(0);
        expect(warnings).toHaveLength(1);
        expect(warnings[0]).toContain("Sam deleted a row over your");
        expect(warnings[0]).toContain("perm do both");
        expect(replicaActor()).toBe("me");
    });

    it("stays silent on a merge that buried nothing", async () => {
        const r = round();
        const sheet = r.sheets.find((s) => s.kind !== "cx")!;
        seedReplica(r, "me");

        let t = 9_000;
        const theirs = applyOp(
            getReplica()!,
            { kind: "cellText", sheetId: sheet.id, col: 0, row: 1, text: "their turn" },
            { actor: "sam", clock: createClock("sam", () => t++) },
        );

        applyRemoteDoc(r, theirs);
        expect(warnings).toEqual([]);
    });
});

describe("the identity a session gives this machine", () => {
    it("is adopted by the replica, so its cells cannot collide with a peer's", async () => {
        const r = round();
        seedReplica(r);
        expect(replicaActor()).toBe("");

        const session = await startForRound(r);
        expect(session).not.toBeNull();
        expect(session!.endpointId).not.toBe("");
        expect(replicaActor()).toBe(session!.endpointId);

        // And a remote apply does not hand it back.
        const sheet = r.sheets.find((s) => s.kind !== "cx")!;
        let t = 9_000;
        applyRemoteDoc(
            r,
            applyOp(
                getReplica()!,
                { kind: "cellText", sheetId: sheet.id, col: 0, row: 1, text: "theirs" },
                { actor: "sam", clock: createClock("sam", () => t++) },
            ),
        );
        expect(replicaActor()).toBe(session!.endpointId);
        await endSession();
    });
});
