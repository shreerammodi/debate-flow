import { configure, render, screen, waitFor } from "@testing-library/react";
import type Handsontable from "handsontable";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import HotGrid from "@/components/flow/HotGrid";
import { LOCK_TTL_MS, type Lock } from "@/lib/collab/presence";
import { getActiveHot } from "@/lib/grid/hotInstance";
import { setLocks } from "@/lib/grid/lockBridge";
import { LOCK_CLASS } from "@/lib/grid/lockDecor";
import { makeFlowRound } from "@/lib/model/flow";
import { useFlowStore } from "@/lib/store/useFlowStore";

const SAM = "kx7f9q2wsamendpoint";
const KIM = "b3m8t1z4kimendpoint";

const round = makeFlowRound();
const sheetId = round.sheets[0].id;

const held = (endpointId: string, col: number, row: number, heldAt: number): Lock => ({
    endpointId,
    sheetId,
    col,
    row,
    heldAt,
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
    setLocks([]);
});

// Each test mounts a real Handsontable over a 250-row sheet and waits for two
// imperative repaints. That is genuinely slow, and slower still when the whole
// suite runs in parallel on a two-core runner, so this file gets ceilings that
// reflect the work rather than ones tuned to a quiet machine. Both are needed:
// `testTimeout` bounds the test, while `asyncUtilTimeout` bounds a single
// `waitFor`, which otherwise gives up after a second and reports the repaint
// that had not landed yet as a wrong class.
vi.setConfig({ testTimeout: 30_000 });
configure({ asyncUtilTimeout: 15_000 });

describe("HotGrid lock surface", () => {
    it("marks the cell a peer holds and clears it on release", async () => {
        const hot = await mount();

        setLocks([held(SAM, 1, 2, Date.now())]);
        await waitFor(() => expect(hot.getCell(2, 1)).toHaveClass(LOCK_CLASS));
        expect(hot.getCell(2, 0)).not.toHaveClass(LOCK_CLASS);

        setLocks([]);
        await waitFor(() => expect(hot.getCell(2, 1)).not.toHaveClass(LOCK_CLASS));
    });

    it("paints nothing for a lock past its TTL", async () => {
        const hot = await mount();

        setLocks([held(SAM, 1, 2, Date.now() - LOCK_TTL_MS - 1), held(KIM, 0, 4, Date.now())]);

        // Kim's live mark landing is what proves the repaint ran, so Sam's
        // bare cell below is a decision rather than an unrendered pass.
        await waitFor(() => expect(hot.getCell(4, 0)).toHaveClass(LOCK_CLASS));
        expect(hot.getCell(2, 1)).not.toHaveClass(LOCK_CLASS);
    });

    it("answers a refused keystroke with a hint naming the holder", async () => {
        const hot = await mount();
        setLocks([held(SAM, 1, 2, Date.now())]);
        hot.selectCell(2, 1);

        press(hot, "a");

        const hint = await screen.findByTestId("lock-hint");
        expect(hint).toHaveTextContent("Sam is editing this cell");
        expect(hot.getActiveEditor()?.isOpened()).toBeFalsy();
    });

    it("lets a keystroke on a free cell through, silently", async () => {
        const hot = await mount();
        setLocks([held(SAM, 1, 2, Date.now())]);
        hot.selectCell(2, 0);

        // An arrow the grid acts on proves the recorder is receiving these
        // events, so the missing hint below is the guard standing aside
        // rather than a keystroke that never arrived.
        press(hot, "ArrowDown");
        expect(hot.getSelectedRangeLast()?.highlight.row).toBe(3);

        press(hot, "a");
        expect(screen.queryByTestId("lock-hint")).toBeNull();
    });
});
