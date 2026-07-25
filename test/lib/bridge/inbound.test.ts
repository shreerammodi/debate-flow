import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { handleBridgeRequest, resetRevealCycle } from "@/lib/bridge/inbound";
import { setActiveHot } from "@/lib/grid/hotInstance";
import { resetMetaUndo } from "@/lib/grid/metaUndo";
import { makeFlowRound, type CellMeta, type CellSource } from "@/lib/model/flow";
import { useFlowStore } from "@/lib/store/useFlowStore";

/**
 * The slice of Handsontable the inbound routes touch, backed by plain arrays
 * so a write can be asserted on text, decoration and provenance at once.
 */
function makeGrid(rows: number, cols: number) {
    const data: (string | null)[][] = Array.from({ length: rows }, () =>
        Array.from({ length: cols }, () => null),
    );
    const meta = new Map<string, { className?: string; source?: CellSource }>();
    const at = (row: number, col: number) => {
        const key = `${row},${col}`;
        if (!meta.has(key)) meta.set(key, {});
        return meta.get(key)!;
    };
    let selected: [number, number] = [0, 0];

    const hot = {
        countRows: () => data.length,
        countCols: () => cols,
        getDataAtCell: (row: number, col: number) => data[row]?.[col] ?? null,
        setDataAtCell: (changes: [number, number, string | null][]) => {
            for (const [row, col, value] of changes) {
                if (data[row]) data[row][col] = value;
            }
        },
        getCellMeta: (row: number, col: number) => at(row, col),
        setCellMeta: (row: number, col: number, key: string, value: unknown) => {
            if (key === "source") at(row, col).source = value as CellSource | undefined;
            else at(row, col).className = value as string;
        },
        getSelectedLast: () => selected,
        selectCell: (row: number, col: number) => {
            selected = [row, col];
        },
        alter: (_action: string, _index: number, amount: number) => {
            for (let i = 0; i < amount; i++) data.push(Array.from({ length: cols }, () => null));
        },
        render: vi.fn(),
    };
    return { hot, data, at, select: (row: number, col: number) => hot.selectCell(row, col) };
}

const send = (items: unknown[], mode = "column") =>
    handleBridgeRequest("flow", { mode, docTitle: "AT - Cap K", items });

const tag = { kind: "tag", text: "Perm solves", source: "cmsrc1.a", key: "doc-1|perm solves" };
const cite = { kind: "cite", text: "Smith 24", source: "cmsrc1.b", key: "doc-1|smith 24" };
const block = { kind: "block", text: "Cap K", source: "cmsrc1.c", key: "doc-1|cap k" };

function loadRound() {
    useFlowStore.getState().loadRound(makeFlowRound({ role: "aff" }));
}

/** The sheet a write lands in: the one the store made active. */
function activeSheetTitle(): string {
    const state = useFlowStore.getState();
    return state.round!.sheets.find((s) => s.id === state.activeSheetId)!.title;
}

beforeEach(() => {
    // The bridge is desktop-only, so every route needs the shell's global.
    (window as unknown as Record<string, unknown>).__TAURI_INTERNALS__ = {};
    setActiveHot(null, null);
    resetMetaUndo();
    resetRevealCycle();
    useFlowStore.setState({
        round: null,
        activeSheetId: null,
        insertPaste: false,
        revealTarget: null,
        cardmirrorEnabled: true,
    });
});

afterEach(() => {
    delete (window as unknown as Record<string, unknown>).__TAURI_INTERNALS__;
});

describe("the flow route", () => {
    it("writes a row per item downward from the active cell", () => {
        loadRound();
        const grid = makeGrid(10, 3);
        grid.select(2, 1);
        setActiveHot(grid.hot as never, vi.fn());

        const reply = send([block, tag, cite]);

        expect(reply.status).toBe(200);
        expect(reply.body).toMatchObject({ ok: true, written: 2, row: 2, col: 1 });
        expect(grid.data[2][1]).toBe("Cap K");
        expect(grid.data[3][1]).toBe("Perm solves\nSmith 24");
        expect(grid.at(2, 1).className).toBe("flow-bold");
        expect(grid.at(3, 1).className).toBe("flow-card");
        expect(grid.at(3, 1).source).toEqual({
            app: "cardmirror",
            token: "cmsrc1.a",
            key: "doc-1|perm solves",
            title: "AT - Cap K",
        });
    });

    it("leaves the cursor under the send so a second one stacks", () => {
        loadRound();
        const grid = makeGrid(10, 3);
        grid.select(2, 1);
        setActiveHot(grid.hot as never, vi.fn());

        send([block, tag]);
        expect(grid.hot.getSelectedLast()).toEqual([4, 1]);
        send([tag]);
        expect(grid.data[4][1]).toBe("Perm solves");
    });

    it("overwrites the column by default", () => {
        loadRound();
        const grid = makeGrid(10, 3);
        grid.data[0][0] = "old";
        grid.data[1][0] = "keep";
        setActiveHot(grid.hot as never, vi.fn());

        send([tag]);
        expect(grid.data[0][0]).toBe("Perm solves");
        expect(grid.data[1][0]).toBe("keep");
    });

    it("pushes the column down when insert paste is on", () => {
        loadRound();
        useFlowStore.setState({ insertPaste: true });
        const grid = makeGrid(10, 3);
        grid.data[0][0] = "old";
        grid.data[1][0] = "older";
        grid.at(0, 0).className = "flow-bold";
        setActiveHot(grid.hot as never, vi.fn());

        send([tag]);
        expect(grid.data[0][0]).toBe("Perm solves");
        expect(grid.data[1][0]).toBe("old");
        expect(grid.data[2][0]).toBe("older");
        expect(grid.at(1, 0).className).toBe("flow-bold");
        expect(grid.at(0, 0).className).toBe("flow-card");
    });

    it("grows the grid rather than dropping a send off the bottom", () => {
        loadRound();
        const grid = makeGrid(2, 1);
        grid.select(1, 0);
        setActiveHot(grid.hot as never, vi.fn());

        send([block, tag, { kind: "analytic", text: "No link" }]);
        expect(grid.data).toHaveLength(4);
        expect(grid.data[3][0]).toBe("No link");
    });

    it("names the sheet it wrote to and reports the cell count", () => {
        loadRound();
        const grid = makeGrid(10, 1);
        setActiveHot(grid.hot as never, vi.fn());

        expect(send([tag, cite], "cell").body).toMatchObject({
            ok: true,
            written: 1,
            sheet: activeSheetTitle(),
        });
        expect(grid.data[0][0]).toBe("Perm solves\nSmith 24");
    });

    it("answers no-active-sheet with no grid and no-active-cell with no selection", () => {
        expect(send([tag]).body).toEqual({ ok: false, error: "no-active-sheet" });

        loadRound();
        const grid = makeGrid(10, 1);
        setActiveHot({ ...grid.hot, getSelectedLast: () => undefined } as never, vi.fn());
        expect(send([tag]).body).toEqual({ ok: false, error: "no-active-cell" });
    });

    it("rejects a malformed body and an unknown route", () => {
        expect(handleBridgeRequest("flow", { items: [] })).toEqual({
            status: 400,
            body: { ok: false, error: "bad-request" },
        });
        expect(handleBridgeRequest("nonsense", {}).status).toBe(400);
    });
});

describe("the reveal route", () => {
    const sourced = (key: string): CellMeta => ({ source: { app: "cardmirror", token: "t", key } });

    /** Two hits on the CX sheet (order -1, so first) and one on the flow sheet. */
    function roundWithHits() {
        const round = makeFlowRound({ role: "aff" });
        round.sheets[0].meta = { "3,2": sourced("doc-1|perm"), "1,0": sourced("doc-1|perm") };
        round.sheets[1].meta = { "0,1": sourced("doc-1|smith") };
        useFlowStore.getState().loadRound(round);
        return round;
    }

    const reveal = (keys: string[]) => handleBridgeRequest("reveal", { keys });

    it("counts every match, names the sheets, and selects the first", () => {
        const round = roundWithHits();
        const body = reveal(["doc-1|perm", "doc-1|smith"]).body as Record<string, unknown>;

        expect(body.ok).toBe(true);
        expect(body.matches).toBe(3);
        expect(body.sheets).toEqual([round.sheets[0].title, round.sheets[1].title]);
        expect(body).toMatchObject({ sheet: round.sheets[0].title, row: 1, col: 0 });
        expect(useFlowStore.getState().revealTarget).toMatchObject({
            sheetId: round.sheets[0].id,
            row: 1,
            col: 0,
        });
    });

    it("walks to the next match when the same keys come back", () => {
        roundWithHits();
        const keys = ["doc-1|perm", "doc-1|smith"];
        expect(reveal(keys).body).toMatchObject({ row: 1, col: 0 });
        expect(reveal(keys).body).toMatchObject({ row: 3, col: 2 });
        expect(reveal(keys).body).toMatchObject({ row: 0, col: 1 });
        expect(reveal(keys).body).toMatchObject({ row: 1, col: 0 });
    });

    it("restarts at the first match when a different card asks", () => {
        roundWithHits();
        reveal(["doc-1|perm", "doc-1|smith"]);
        expect(reveal(["doc-1|perm"]).body).toMatchObject({ row: 1, col: 0 });
    });

    it("reports zero matches without selecting anything", () => {
        roundWithHits();
        expect(reveal(["doc-1|absent"]).body).toEqual({ ok: true, matches: 0 });
        expect(useFlowStore.getState().revealTarget).toBeNull();
    });

    it("answers no-round with nothing open and rejects an empty key list", () => {
        expect(reveal(["doc-1|perm"]).body).toEqual({ ok: false, error: "no-round" });
        loadRound();
        expect(handleBridgeRequest("reveal", { keys: [] }).status).toBe(400);
    });
});

describe("the desktop-only gate", () => {
    it("turns every route away when the switch is off, without touching the grid", () => {
        loadRound();
        const grid = makeGrid(10, 3);
        setActiveHot(grid.hot as never, vi.fn());
        useFlowStore.setState({ cardmirrorEnabled: false });

        expect(send([tag]).body).toEqual({ ok: false, error: "integration-disabled" });
        expect(handleBridgeRequest("reveal", { keys: ["doc-1|perm"] }).body).toEqual({
            ok: false,
            error: "integration-disabled",
        });
        expect(grid.data[0][0]).toBeNull();
    });

    it("turns every route away on the web build", () => {
        loadRound();
        const grid = makeGrid(10, 3);
        setActiveHot(grid.hot as never, vi.fn());
        delete (window as unknown as Record<string, unknown>).__TAURI_INTERNALS__;

        expect(send([tag]).body).toEqual({ ok: false, error: "integration-disabled" });
        expect(grid.data[0][0]).toBeNull();
    });
});
