import { render, waitFor } from "@testing-library/react";
import type Handsontable from "handsontable";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import HotGrid from "@/components/flow/HotGrid";
import { getActiveHot } from "@/lib/grid/hotInstance";
import { makeFlowRound } from "@/lib/model/flow";
import { useFlowStore } from "@/lib/store/useFlowStore";

const round = makeFlowRound();
const sheetId = round.sheets[0].id;

async function mount() {
    render(<HotGrid sheetId={sheetId} pane={1} />);
    await waitFor(() => expect(getActiveHot()).not.toBeNull());
    return getActiveHot()!;
}

/**
 * A real keydown through Handsontable's own recorder, so beforeKeyDown runs.
 * The legacy keyCode rides along because a fast edit is exactly what a missing
 * one suppresses: Handsontable reads keyCode 0 as a function key and leaves the
 * cell alone, which no browser ever sends for a printable key.
 */
function press(hot: Handsontable, key: string) {
    const target = document.activeElement ?? hot.rootElement;
    target.dispatchEvent(
        new KeyboardEvent("keydown", {
            key,
            keyCode: key.length === 1 ? key.toUpperCase().charCodeAt(0) : 0,
            bubbles: true,
            cancelable: true,
        }),
    );
}

function editorInput(hot: Handsontable): HTMLTextAreaElement {
    const input = hot.rootElement.querySelector<HTMLTextAreaElement>("textarea.handsontableInput");
    expect(input).not.toBeNull();
    return input!;
}

/**
 * jsdom runs no default action for a keydown, so the character Handsontable
 * leaves to the browser never arrives. What the editor was handed is the part
 * this side decides; typing into it is the browser's.
 */
function typeInto(input: HTMLTextAreaElement, char: string) {
    const at = input.selectionStart;
    input.value = `${input.value.slice(0, at)}${char}${input.value.slice(input.selectionEnd)}`;
}

beforeEach(() => {
    useFlowStore.setState({ round, activeSheetId: sheetId, splitSheetId: null, appendEdit: true });
});

afterEach(() => {
    useFlowStore.setState({ appendEdit: true });
});

// Each test mounts a real Handsontable over a 250-row sheet, which is genuinely
// slow, and slower still when the whole suite runs in parallel.
vi.setConfig({ testTimeout: 30_000 });

describe("what a printable key does to a cell that already has text", () => {
    it("keeps the text and leaves the caret past its end", async () => {
        const hot = await mount();
        hot.setDataAtCell(2, 0, "perm");
        hot.selectCell(2, 0);

        press(hot, "x");

        const input = editorInput(hot);
        expect(hot.getActiveEditor()?.isOpened()).toBe(true);
        expect(input.value).toBe("perm");
        expect(input.selectionStart).toBe(4);

        typeInto(input, "x");
        press(hot, "Enter");

        expect(hot.getDataAtCell(2, 0)).toBe("permx");
    });

    it("writes over the cell with append mode off", async () => {
        useFlowStore.setState({ appendEdit: false });
        const hot = await mount();
        hot.setDataAtCell(2, 0, "perm");
        hot.selectCell(2, 0);

        press(hot, "x");

        const input = editorInput(hot);
        expect(hot.getActiveEditor()?.isOpened()).toBe(true);
        expect(input.value).toBe("");

        typeInto(input, "x");
        press(hot, "Enter");

        expect(hot.getDataAtCell(2, 0)).toBe("x");
    });

    // F2 opens in full edit mode, where Handsontable seeds the box itself and
    // parks the caret at the end - so append mode has nothing to add, and must
    // not double the text by adding it anyway.
    it("leaves an editor opened for a full edit alone", async () => {
        const hot = await mount();
        hot.setDataAtCell(3, 1, "perm");
        hot.selectCell(3, 1);

        press(hot, "F2");

        expect(editorInput(hot).value).toBe("perm");
    });
});
