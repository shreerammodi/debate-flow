import { render, waitFor } from "@testing-library/react";
import type Handsontable from "handsontable";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import HotGrid from "@/components/flow/HotGrid";
import { getActiveHot } from "@/lib/grid/hotInstance";
import {
    editingHere,
    type HeldCell,
    setClaimHandler,
    setCursorHandler,
    setPresences,
} from "@/lib/grid/presenceBridge";
import { makeFlowRound } from "@/lib/model/flow";
import { useCollabStore } from "@/lib/store/useCollabStore";
import { useFlowStore } from "@/lib/store/useFlowStore";

const round = makeFlowRound();
const sheetId = round.sheets[0].id;

const claims: HeldCell[] = [];
const cursors: HeldCell[] = [];

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
    claims.length = 0;
    cursors.length = 0;
    setClaimHandler((cell) => claims.push(cell));
    setCursorHandler((cell) => cursors.push(cell));
    useFlowStore.setState({ round, activeSheetId: sheetId, splitSheetId: null });
});

afterEach(() => {
    setClaimHandler(null);
    setCursorHandler(null);
    setPresences([]);
    useCollabStore.getState().reset();
});

// Each test mounts a real Handsontable over a 250-row sheet, which is genuinely
// slow, and slower still when the whole suite runs in parallel.
vi.setConfig({ testTimeout: 30_000 });

describe("what this side tells a partner about where it is", () => {
    it("claims the cell an editor opens on", async () => {
        const hot = await mount();

        hot.selectCell(2, 0);
        press(hot, "F2");

        expect(hot.getActiveEditor()?.isOpened()).toBe(true);
        expect(claims.at(-1)).toMatchObject({ col: 0, row: 2 });
        expect(editingHere()).toBe(true);
    });

    // The claim is released by being asked about, not by an event, because
    // Handsontable fires nothing when an editor closes. Each way a debater
    // leaves a cell is held separately: Enter moves the selection, Escape
    // moves nothing at all, and neither one announces the close.
    it("stops backing the claim once the editor commits", async () => {
        const hot = await mount();
        hot.selectCell(2, 0);
        press(hot, "F2");
        expect(editingHere()).toBe(true);

        press(hot, "Enter");

        expect(hot.getActiveEditor()?.isOpened()).toBeFalsy();
        expect(editingHere()).toBe(false);
    });

    it("stops backing the claim when the editor is abandoned instead", async () => {
        const hot = await mount();
        hot.selectCell(3, 1);
        press(hot, "F2");
        expect(editingHere()).toBe(true);

        press(hot, "Escape");

        expect(hot.getActiveEditor()?.isOpened()).toBeFalsy();
        expect(editingHere()).toBe(false);
    });

    it("keeps reporting the cursor after an edit, not only on the next one", async () => {
        const hot = await mount();
        hot.selectCell(2, 0);
        press(hot, "F2");
        press(hot, "Enter");

        cursors.length = 0;
        hot.selectCell(5, 0);
        hot.selectCell(7, 0);

        expect(cursors.at(-1)).toMatchObject({ col: 0, row: 7 });
    });

    it("answers for no editor when no grid is mounted", () => {
        expect(editingHere()).toBe(false);
    });
});
