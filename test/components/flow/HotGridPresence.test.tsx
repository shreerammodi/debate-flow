import { render, screen, waitFor } from "@testing-library/react";
import type Handsontable from "handsontable";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import HotGrid from "@/components/flow/HotGrid";
import { PRESENCE_TTL_MS, type Presence } from "@/lib/collab/presence";
import { modelCol } from "@/lib/grid/colSpace";
import { getActiveHot } from "@/lib/grid/hotInstance";
import { setPresences } from "@/lib/grid/presenceBridge";
import { LOCK_CLASS, PEER_CLASS } from "@/lib/grid/presenceDecor";
import { makeFlowRound } from "@/lib/model/flow";
import { useCollabStore } from "@/lib/store/useCollabStore";
import { useFlowStore } from "@/lib/store/useFlowStore";

const SAM = "kx7f9q2wsamendpoint";
const KIM = "b3m8t1z4kimendpoint";

const round = makeFlowRound();
const sheetId = round.sheets[0].id;

const at = (
    endpointId: string,
    col: number,
    row: number,
    heldAt: number,
    editing: boolean,
): Presence => ({
    endpointId,
    sheetId,
    col: modelCol(col),
    row,
    heldAt,
    editing,
});

/** The pane, mounted and past its first imperative data load. */
async function mount() {
    render(<HotGrid sheetId={sheetId} pane={1} />);
    await waitFor(() => expect(getActiveHot()).not.toBeNull());
    return getActiveHot()!;
}

/** A real keydown through Handsontable's own recorder, so beforeKeyDown runs. */
function press(hot: Handsontable, key: string) {
    const target = document.activeElement ?? hot.rootElement;
    target.dispatchEvent(new KeyboardEvent("keydown", { key, bubbles: true, cancelable: true }));
}

beforeEach(() => {
    // Presence carries a wall-clock `heldAt` and `expire` drops it one second
    // later, while mounting a real Handsontable over a 250-row sheet can
    // outlast that on a slow machine: the cell then paints bare because the
    // entry died, not because the decoration is wrong. A peer refreshes every
    // HEARTBEAT_MS in a real session, so freezing the clock is what a live one
    // looks like. Only Date is faked; the grid's own timers and the repaints
    // they drive have to keep running.
    vi.useFakeTimers({ toFake: ["Date"] });
    useFlowStore.setState({
        round,
        activeSheetId: sheetId,
        splitSheetId: null,
        contacts: {
            [SAM]: { name: "Sam", role: "partner" },
            [KIM]: { name: "Kim", role: "coach" },
        },
    });
});

afterEach(() => {
    vi.useRealTimers();
    setPresences([]);
    useCollabStore.getState().reset();
});

// Each test mounts a real Handsontable over a 250-row sheet and waits for two
// imperative repaints. That is genuinely slow, and slower still when the whole
// suite runs in parallel, so this file gets a ceiling that reflects the work
// rather than one tuned to a quiet machine.
vi.setConfig({ testTimeout: 30_000 });

describe("HotGrid presence surface", () => {
    it("marks the cell a peer holds and clears it on release", async () => {
        const hot = await mount();

        setPresences([at(SAM, 1, 2, Date.now(), true)]);
        await waitFor(() => expect(hot.getCell(2, 1)).toHaveClass(LOCK_CLASS));
        expect(hot.getCell(2, 0)).not.toHaveClass(LOCK_CLASS);

        setPresences([]);
        await waitFor(() => expect(hot.getCell(2, 1)).not.toHaveClass(LOCK_CLASS));
    });

    it("paints nothing for presence past its TTL", async () => {
        const hot = await mount();

        setPresences([
            at(SAM, 1, 2, Date.now() - PRESENCE_TTL_MS - 1, true),
            at(KIM, 0, 4, Date.now(), true),
        ]);

        // Kim's live mark landing is what proves the repaint ran, so Sam's
        // bare cell below is a decision rather than an unrendered pass.
        await waitFor(() => expect(hot.getCell(4, 0)).toHaveClass(LOCK_CLASS));
        expect(hot.getCell(2, 1)).not.toHaveClass(LOCK_CLASS);
    });

    it("answers a refused keystroke with a hint naming the holder", async () => {
        const hot = await mount();
        setPresences([at(SAM, 1, 2, Date.now(), true)]);
        hot.selectCell(2, 1);

        press(hot, "a");

        const hint = await screen.findByTestId("lock-hint");
        expect(hint).toHaveTextContent("Sam is editing this cell");
        expect(hot.getActiveEditor()?.isOpened()).toBeFalsy();
    });

    it("lets a keystroke on a free cell through, silently", async () => {
        const hot = await mount();
        setPresences([at(SAM, 1, 2, Date.now(), true)]);
        hot.selectCell(2, 0);

        // An arrow the grid acts on proves the recorder is receiving these
        // events, so the missing hint below is the guard standing aside
        // rather than a keystroke that never arrived.
        press(hot, "ArrowDown");
        expect(hot.getSelectedRangeLast()?.highlight.row).toBe(3);

        press(hot, "a");
        expect(screen.queryByTestId("lock-hint")).toBeNull();
    });

    it("shows a parked cursor without blocking the cell", async () => {
        const hot = await mount();
        setPresences([at(SAM, 1, 2, Date.now(), false)]);

        await waitFor(() => expect(hot.getCell(2, 1)).toHaveClass(PEER_CLASS));
        expect(hot.getCell(2, 1)).not.toHaveClass(LOCK_CLASS);

        hot.selectCell(2, 1);
        press(hot, "a");
        expect(screen.queryByTestId("lock-hint")).toBeNull();

        // An arrow the grid acts on proves the recorder is receiving these
        // events, so the missing hint above is the guard standing aside
        // rather than a keystroke that never arrived.
        press(hot, "ArrowDown");
        expect(hot.getSelectedRangeLast()?.highlight.row).toBe(3);

        // F2 runs the same refusal path a printable key does, and reaches the
        // editor when the guard lets it by, so an open editor on the parked
        // cell is the cell accepting input rather than merely staying quiet.
        hot.selectCell(2, 1);
        press(hot, "F2");
        expect(screen.queryByTestId("lock-hint")).toBeNull();
        expect(hot.getActiveEditor()?.isOpened()).toBe(true);
        hot.getActiveEditor()?.close();
    });

    it("badges the cell with the peer's initial and drops it when they move", async () => {
        const hot = await mount();

        setPresences([at(SAM, 1, 2, Date.now(), false)]);
        await waitFor(() => expect(hot.getCell(2, 1)?.dataset.peer).toBe("S"));

        setPresences([at(SAM, 1, 5, Date.now(), false)]);
        await waitFor(() => expect(hot.getCell(5, 1)?.dataset.peer).toBe("S"));
        expect(hot.getCell(2, 1)?.dataset.peer).toBeUndefined();
    });

    it("refuses a coach's keystroke instead of losing it", async () => {
        useCollabStore.setState({ selfRole: "coach" });
        const hot = await mount();
        hot.selectCell(2, 0);

        // Navigation still works, so a cell that stays empty below is the
        // grid refusing the write rather than an event that never arrived.
        press(hot, "ArrowDown");
        expect(hot.getSelectedRangeLast()?.highlight.row).toBe(3);

        hot.selectCell(2, 0);
        press(hot, "a");
        expect(hot.getDataAtCell(2, 0)).toBeFalsy();

        // F2 opens the editor on a writable cell, so a shut editor here is the
        // read-only grid and not a key the recorder swallowed.
        press(hot, "F2");
        expect(hot.getActiveEditor()?.isOpened()).toBeFalsy();
    });
});
