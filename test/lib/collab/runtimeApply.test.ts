import { beforeEach, describe, expect, it, vi } from "vitest";

vi.mock("sonner", () => ({
    toast: Object.assign(() => {}, {
        warning: () => {},
        error: () => {},
        success: () => {},
        info: () => {},
    }),
}));

import { seedDoc } from "@/lib/collab/doc";
import { applyOp, type OpContext } from "@/lib/collab/ops";
import { getReplica } from "@/lib/collab/replica";
import { applyRemoteDoc, endSession } from "@/lib/collab/runtime";
import { createClock } from "@/lib/collab/stamp";
import type { CollabDoc } from "@/lib/collab/types";
import { makeFlowRound, type FlowRound } from "@/lib/model/flow";
import { useFlowStore } from "@/lib/store/useFlowStore";

function round(): FlowRound {
    const r = makeFlowRound({});
    const sheet = r.sheets.find((s) => s.kind !== "cx")!;
    sheet.data = [
        ["perm", "link"],
        ["cap bad", "turn"],
    ];
    return r;
}

/** A partner's replica of the same file, plus one write of their own. */
function peerDoc(base: FlowRound, apply: (doc: CollabDoc, ctx: OpContext) => CollabDoc): CollabDoc {
    let t = 5_000;
    const ctx: OpContext = { actor: "sam", clock: createClock("sam", () => t++) };
    return apply(seedDoc(base), ctx);
}

let open: FlowRound;
let sheetId: string;

beforeEach(async () => {
    await endSession();
    useFlowStore.setState({ collabEnabled: true, collabRelayEnabled: true, shadowMode: false });
    open = round();
    sheetId = open.sheets.find((s) => s.kind !== "cx")!.id;
    useFlowStore.getState().loadRound(open);
});

function sheetData(): (string | null)[][] {
    return useFlowStore.getState().round!.sheets.find((s) => s.id === sheetId)!.data;
}

describe("a partner's change reaching this machine", () => {
    it("shows up in the open round, not only in the replica", () => {
        applyRemoteDoc(
            open,
            peerDoc(open, (doc, ctx) =>
                applyOp(
                    doc,
                    { kind: "cellText", sheetId, col: 0, row: 0, text: "perm, then CP" },
                    ctx,
                ),
            ),
        );

        expect(getReplica()!.sheets[sheetId]).toBeDefined();
        expect(sheetData()[0][0]).toBe("perm, then CP");
    });

    it("bumps updatedAt, so the autosave writes the partner's text to disk", () => {
        const before = useFlowStore.getState().round!.updatedAt;
        applyRemoteDoc(
            open,
            peerDoc(open, (doc, ctx) =>
                applyOp(doc, { kind: "cellText", sheetId, col: 1, row: 1, text: "no turn" }, ctx),
            ),
        );
        expect(useFlowStore.getState().round!.updatedAt).toBeGreaterThanOrEqual(before);
        expect(sheetData()[1][1]).toBe("no turn");
    });

    it("carries a partner's row insert into the round", () => {
        applyRemoteDoc(
            open,
            peerDoc(open, (doc, ctx) => {
                const grown = applyOp(doc, { kind: "insertCell", sheetId, col: 0, row: 1 }, ctx);
                return applyOp(
                    grown,
                    { kind: "cellText", sheetId, col: 0, row: 1, text: "extend" },
                    ctx,
                );
            }),
        );
        expect(sheetData().map((r) => r[0])).toEqual(["perm", "extend", "cap bad"]);
    });

    it("leaves the round alone when a partner sends nothing new", () => {
        const before = useFlowStore.getState().round;
        applyRemoteDoc(open, seedDoc(open));
        expect(sheetData()).toEqual(before!.sheets.find((s) => s.id === sheetId)!.data);
    });

    it("moves off a sheet a partner deleted out from under the cursor", () => {
        useFlowStore.setState({ activeSheetId: sheetId });
        applyRemoteDoc(
            open,
            peerDoc(open, (doc, ctx) => applyOp(doc, { kind: "removeSheet", sheetId }, ctx)),
        );
        const state = useFlowStore.getState();
        expect(state.round!.sheets.some((s) => s.id === sheetId)).toBe(false);
        expect(state.activeSheetId).not.toBe(sheetId);
    });
});
