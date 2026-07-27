import { beforeEach, describe, expect, it } from "vitest";

import { projectDoc } from "@/lib/collab/doc";
import { isReplicatedSource, rowOpFromHook, textOpsFromChanges } from "@/lib/collab/gridOps";
import { persistReplica, recoverReplica } from "@/lib/collab/persist";
import { clearReplica, driftedSheetIds, getReplica, recordOp } from "@/lib/collab/replica";
import { setSidecarFs, type SidecarFs } from "@/lib/collab/sidecarFs";
import type { GridChange } from "@/lib/grid/staleSource";
import { makeFlowRound, type FlowRound } from "@/lib/model/flow";
import { serializeFlow } from "@/lib/persistence/flowFile";
import { useFlowStore } from "@/lib/store/useFlowStore";

interface FakeSidecarFs extends SidecarFs {
    files: Map<string, string>;
}

function fakeFs(): FakeSidecarFs {
    const files = new Map<string, string>();
    return {
        files,
        async read(id) {
            return files.get(id) ?? null;
        },
        async write(id, text) {
            files.set(id, text);
        },
    };
}

/**
 * Drives the replica through the same functions the grid seams call, without
 * mounting a grid. `afterChange` hands its payload to `textOpsFromChanges`
 * once the source passes; the row hooks hand theirs to `rowOpFromHook`.
 */
function typeInto(sheetId: string, changes: GridChange[]): void {
    expect(isReplicatedSource("edit")).toBe(true);
    for (const op of textOpsFromChanges(sheetId, changes)) recordOp(op);
}

let fs: FakeSidecarFs;

beforeEach(() => {
    fs = fakeFs();
    setSidecarFs(fs);
    clearReplica();
    useFlowStore.setState({ collabEnabled: true, round: null, activeSheetId: null });
});

function openedRound(): { round: FlowRound; sheetId: string } {
    const round = makeFlowRound({});
    const flow = round.sheets.find((s) => s.kind !== "cx")!;
    flow.data = [
        ["perm do both", "no link"],
        ["cap bad", "turn"],
    ];
    return { round, sheetId: flow.id };
}

describe("a round of editing survives a restart", () => {
    it("recovers the same document from the sidecar it wrote", async () => {
        const { round, sheetId } = openedRound();
        await recoverReplica(round, serializeFlow(round));

        // A speech: typed cells, a row inserted above, one more cell.
        typeInto(sheetId, [[0, 0, "perm do both", "perm do both, then CP"]]);
        for (const op of rowOpFromHook("insert", sheetId, 1, 1, undefined)) recordOp(op);
        typeInto(sheetId, [[1, 1, null, "extend Smith"]]);
        recordOp({ kind: "cellMeta", sheetId, col: 0, row: 0, meta: { bold: true } });

        // The store follows, the way the grid snapshot makes it.
        const edited = projectDoc(getReplica()!, round);
        useFlowStore.getState().loadRound(edited);
        // loadRound re-seeds, so put the edited replica back before saving.
        await recoverReplica(edited, serializeFlow(edited));
        const beforeRestart = getReplica();

        await persistReplica(edited, serializeFlow(edited));
        expect(fs.files.has(edited.id)).toBe(true);

        // The app quits and comes back to the same file.
        clearReplica();
        expect(getReplica()).toBeNull();
        await recoverReplica(edited, serializeFlow(edited));

        expect(getReplica()).toEqual(beforeRestart);
        const after = projectDoc(getReplica()!, edited);
        const sheet = after.sheets.find((s) => s.id === sheetId)!;
        expect(sheet.data).toEqual([
            ["perm do both, then CP", "no link"],
            [null, "extend Smith"],
            ["cap bad", "turn"],
        ]);
        expect(sheet.meta["0,0"]).toEqual({ bold: true });
        expect(driftedSheetIds(after)).toEqual([]);
    });

    it("falls back to the file when the flow changed outside ebb", async () => {
        const { round, sheetId } = openedRound();
        await recoverReplica(round, serializeFlow(round));
        typeInto(sheetId, [[0, 0, "perm do both", "typed in this session"]]);
        await persistReplica(round, serializeFlow(round));

        // Someone edited the .ebb while ebb was closed, so the hash no longer
        // names the sidecar's flow.
        const changed: FlowRound = {
            ...round,
            sheets: round.sheets.map((s) =>
                s.id === sheetId ? { ...s, data: [["edited elsewhere", "no link"]] } : s,
            ),
        };
        clearReplica();
        await recoverReplica(changed, serializeFlow(changed));

        const sheet = projectDoc(getReplica()!, changed).sheets.find((s) => s.id === sheetId)!;
        expect(sheet.data[0][0]).toBe("edited elsewhere");
        expect(driftedSheetIds(changed)).toEqual([]);
    });

    it("repairs a sheet whose hook never fired, at the next save", async () => {
        const { round, sheetId } = openedRound();
        await recoverReplica(round, serializeFlow(round));

        // A write that reached the store and skipped the replica entirely.
        const missed: FlowRound = {
            ...round,
            sheets: round.sheets.map((s) =>
                s.id === sheetId ? { ...s, data: [["perm do both", "NEVER REPORTED"]] } : s,
            ),
        };
        expect(driftedSheetIds(missed)).toEqual([sheetId]);

        await persistReplica(missed, serializeFlow(missed));
        expect(driftedSheetIds(missed)).toEqual([]);
        const sheet = projectDoc(getReplica()!, missed).sheets.find((s) => s.id === sheetId)!;
        expect(sheet.data[0][1]).toBe("NEVER REPORTED");
    });

    it("writes no sidecar for a debater who never switched shared editing on", async () => {
        const { round, sheetId } = openedRound();
        useFlowStore.setState({ collabEnabled: false });
        await recoverReplica(round, serializeFlow(round));

        // The replica still tracks every edit: one code path, not two.
        typeInto(sheetId, [[0, 0, "perm do both", "still replicated"]]);
        await persistReplica(round, serializeFlow(round));

        const sheet = projectDoc(getReplica()!, round).sheets.find((s) => s.id === sheetId)!;
        expect(sheet.data[0][0]).toBe("still replicated");
        expect(fs.files.size).toBe(0);
    });
});
