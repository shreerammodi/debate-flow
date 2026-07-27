import { render, waitFor } from "@testing-library/react";
import { beforeEach, describe, expect, it, vi } from "vitest";

vi.mock("sonner", () => ({
    toast: Object.assign(() => {}, {
        warning: () => {},
        error: () => {},
        success: () => {},
        info: () => {},
    }),
}));

import HotGrid from "@/components/flow/HotGrid";
import { seedDoc } from "@/lib/collab/doc";
import { applyOp, type CollabOp, type OpContext } from "@/lib/collab/ops";
import { getReplica } from "@/lib/collab/replica";
import { applyRemoteDoc } from "@/lib/collab/runtime";
import { createClock } from "@/lib/collab/stamp";
import type { CollabDoc } from "@/lib/collab/types";
import { getActiveHot } from "@/lib/grid/hotInstance";
import { makeFlowRound, type FlowRound } from "@/lib/model/flow";
import { useFlowStore } from "@/lib/store/useFlowStore";

// A real Handsontable over a 250-row sheet, twice per test. Slow on purpose,
// and slower under a parallel suite, so the ceiling reflects the work.
vi.setConfig({ testTimeout: 30_000 });

let round: FlowRound;
let sheetId: string;

/** What the partner sends: their replica of the same file, plus their writes. */
function fromSam(...ops: CollabOp[]): CollabDoc {
    let t = 5_000;
    const ctx: OpContext = { actor: "sam", clock: createClock("sam", () => t++) };
    let doc = seedDoc(round);
    for (const op of ops) doc = applyOp(doc, op, ctx);
    return doc;
}

async function mount() {
    render(<HotGrid sheetId={sheetId} pane={1} />);
    await waitFor(() => expect(getActiveHot()).not.toBeNull());
    const hot = getActiveHot()!;
    await waitFor(() => expect(hot.getDataAtCell(0, 0)).toBe("perm"));
    return hot;
}

beforeEach(() => {
    round = makeFlowRound({});
    const sheet = round.sheets.find((s) => s.kind !== "cx")!;
    sheet.data = [
        ["perm", "link"],
        ["cap bad", "turn"],
    ];
    sheetId = sheet.id;
    useFlowStore.setState({ collabEnabled: true, shadowMode: false });
    useFlowStore.getState().loadRound(round, { activeSheetId: sheetId });
});

describe("a partner's change on the live grid", () => {
    it("shows their text without being asked to reload", async () => {
        const hot = await mount();
        applyRemoteDoc(
            round,
            fromSam({ kind: "cellText", sheetId, col: 1, row: 1, text: "no turn" }),
        );
        await waitFor(() => expect(hot.getDataAtCell(1, 1)).toBe("no turn"));
        expect(hot.getDataAtCell(0, 0)).toBe("perm");
    });

    it("does not push the grid back over the replica it came from", async () => {
        const hot = await mount();
        applyRemoteDoc(
            round,
            fromSam({ kind: "cellText", sheetId, col: 0, row: 0, text: "theirs" }),
        );
        await waitFor(() => expect(hot.getDataAtCell(0, 0)).toBe("theirs"));

        const sheet = useFlowStore.getState().round!.sheets.find((s) => s.id === sheetId)!;
        expect(sheet.data[0][0]).toBe("theirs");
        // The write must not have been recorded as this machine's own edit, or
        // the partner would get their own text back stamped by us.
        const cells = Object.values(getReplica()!.sheets[sheetId].cells);
        expect(cells.find((c) => c.text === "theirs")!.textStamp.actor).toBe("sam");
    });

    it("carries a partner's row insert down the column", async () => {
        const hot = await mount();
        applyRemoteDoc(
            round,
            fromSam(
                { kind: "insertCell", sheetId, col: 0, row: 0 },
                { kind: "cellText", sheetId, col: 0, row: 0, text: "extend" },
            ),
        );
        await waitFor(() => expect(hot.getDataAtCell(0, 0)).toBe("extend"));
        expect(hot.getDataAtCell(1, 0)).toBe("perm");
        expect(hot.getDataAtCell(2, 0)).toBe("cap bad");
        // The neighbouring column never moved.
        expect(hot.getDataAtCell(0, 1)).toBe("link");
    });

    it("leaves the cell under an open editor alone", async () => {
        const hot = await mount();
        hot.selectCell(0, 0);
        hot.getActiveEditor()!.beginEditing();
        hot.getActiveEditor()!.setValue("mine, still typing");

        applyRemoteDoc(
            round,
            fromSam(
                { kind: "cellText", sheetId, col: 0, row: 0, text: "theirs" },
                { kind: "cellText", sheetId, col: 0, row: 1, text: "also theirs" },
            ),
        );
        await waitFor(() => expect(hot.getDataAtCell(1, 0)).toBe("also theirs"));
        expect(hot.getActiveEditor()!.getValue()).toBe("mine, still typing");
    });
});
