import { beforeEach, describe, expect, it, vi } from "vitest";

import { projectDoc } from "@/lib/collab/doc";
import { clearReplica, getReplica } from "@/lib/collab/replica";
import { executeCommand } from "@/lib/commands/commands";
import { setActiveHot } from "@/lib/grid/hotInstance";
import { makeFlowRound, type FlowRound } from "@/lib/model/flow";
import { useFlowStore } from "@/lib/store/useFlowStore";

/**
 * The commands write through the live grid, then report to the replica from
 * the store's snapshot. These fakes stand in for the grid; the store and the
 * replica are the real ones.
 */
function metaGrid() {
    const meta = new Map<string, { className?: string }>();
    const at = (r: number, c: number) => {
        const key = `${r},${c}`;
        if (!meta.has(key)) meta.set(key, {});
        return meta.get(key)!;
    };
    return {
        getSelectedRange: () => [
            {
                highlight: { row: 0, col: 0 },
                getTopLeftCorner: () => ({ row: 0, col: 0 }),
                getBottomRightCorner: () => ({ row: 1, col: 0 }),
            },
        ],
        getCellMeta: (r: number, c: number) => at(r, c),
        setCellMeta: (r: number, c: number, _k: string, v: string) => {
            at(r, c).className = v;
        },
        render: vi.fn(),
    };
}

let round: FlowRound;
let sheetId: string;

function openRound(): void {
    round = makeFlowRound({});
    const flow = round.sheets.find((s) => s.kind !== "cx")!;
    sheetId = flow.id;
    flow.data = [
        ["perm", "link"],
        ["cap bad", "turn"],
        ["extend", null],
    ];
    useFlowStore.getState().loadRound(round);
}

beforeEach(() => {
    clearReplica();
    setActiveHot(null, null, null);
    useFlowStore.setState({ round: null, activeSheetId: null, splitSheetId: null });
});

describe("a decoration toggle reaches the replica", () => {
    it("records the meta of every cell it flipped", () => {
        openRound();
        const grid = metaGrid();
        // The store is the source the op reads, so it carries the new class.
        const mutated = () => {
            useFlowStore
                .getState()
                .updateSheetData(sheetId, round.sheets.find((s) => s.id === sheetId)!.data, {
                    "0,0": { bold: true },
                    "1,0": { bold: true },
                });
        };
        setActiveHot(grid as never, mutated, sheetId);

        executeCommand("format.toggleBold");

        const sheet = projectDoc(getReplica()!, round).sheets.find((s) => s.id === sheetId)!;
        expect(sheet.meta["0,0"]).toEqual({ bold: true });
        expect(sheet.meta["1,0"]).toEqual({ bold: true });
    });

    it("records nothing when no grid names a sheet", () => {
        openRound();
        setActiveHot(metaGrid() as never, vi.fn(), null);
        const before = getReplica();
        executeCommand("format.toggleBold");
        expect(getReplica()).toBe(before);
    });
});

describe("a cell insert reaches the replica as one shift", () => {
    it("opens a rank in its own column and leaves the neighbour alone", () => {
        openRound();
        const data = [
            ["perm", "link"],
            ["cap bad", "turn"],
            ["extend", null],
        ];
        const grid = {
            getSelectedLast: () => [1, 0, 1, 0],
            countRows: () => data.length,
            countCols: () => 2,
            getDataAtCell: (r: number, c: number) => data[r]?.[c] ?? null,
            getCellMeta: () => ({}),
            setCellMeta: vi.fn(),
            setDataAtCell: vi.fn(),
            render: vi.fn(),
        };
        setActiveHot(grid as never, vi.fn(), sheetId);

        executeCommand("cell.insert");

        const sheet = projectDoc(getReplica()!, round).sheets.find((s) => s.id === sheetId)!;
        // One blank opened at row 1 of column 0; column 1 never moved.
        expect(sheet.data.map((r) => r[0])).toEqual(["perm", null, "cap bad", "extend"]);
        expect(sheet.data.map((r) => r[1])).toEqual(["link", "turn", null, null]);
    });
});
