import { render } from "@testing-library/react";
import { act } from "react";
import { describe, it, expect, beforeEach, vi } from "vitest";

import type * as CommandsModule from "@/lib/commands/commands";
import { useKeymap } from "@/lib/keymap/useKeymap";
import { makeFlowRound } from "@/lib/model/flow";
import { isMacPlatform } from "@/lib/platform";
import { useFlowStore } from "@/lib/store/useFlowStore";

// Spy over the real executeCommand: the store-observing tests below need its
// behavior, the scope tests need only to see whether it was reached.
vi.mock("@/lib/commands/commands", async (importOriginal) => {
    const actual = await importOriginal<typeof CommandsModule>();
    return { ...actual, executeCommand: vi.fn(actual.executeCommand) };
});
import { executeCommand } from "@/lib/commands/commands";

const MOD = isMacPlatform() ? { metaKey: true } : { ctrlKey: true };

function Harness() {
    useKeymap();
    return <div data-testid="harness" />;
}

function dispatchKey(key: string, init: Partial<KeyboardEventInit> = {}, target?: EventTarget) {
    act(() => {
        const event = new KeyboardEvent("keydown", {
            key,
            bubbles: true,
            cancelable: true,
            ...init,
        });
        (target ?? window).dispatchEvent(event);
    });
}

function freshRound() {
    const round = makeFlowRound({});
    useFlowStore.getState().loadRound(round);
    useFlowStore.getState().addSheet({ title: "DA", group: "neg" });
}

describe("useKeymap", () => {
    beforeEach(() => {
        vi.clearAllMocks();
        document.body.innerHTML = "";
        useFlowStore.setState({
            round: null,
            activeSheetId: null,
            keymapOverrides: {},
            cheatsheetOpen: false,
        });
    });

    it("fires bare-key sheet chords outside text entry", () => {
        freshRound();
        const state = () => useFlowStore.getState();
        const second = state().activeSheetId!;
        render(<Harness />);
        dispatchKey("[");
        expect(state().activeSheetId).not.toBe(second);
        dispatchKey("]");
        expect(state().activeSheetId).toBe(second);
    });

    it("toggles the cheatsheet on ? and respects user overrides", () => {
        freshRound();
        render(<Harness />);
        dispatchKey("?", { shiftKey: true });
        expect(useFlowStore.getState().cheatsheetOpen).toBe(true);

        useFlowStore.setState({ keymapOverrides: { "help.open": "F1" } });
        dispatchKey("F1");
        expect(useFlowStore.getState().cheatsheetOpen).toBe(false);
        // The old chord no longer fires after the override.
        dispatchKey("?", { shiftKey: true });
        expect(useFlowStore.getState().cheatsheetOpen).toBe(false);
    });

    it("does not fire bare-key chords while typing in a text field", () => {
        freshRound();
        const state = () => useFlowStore.getState();
        const second = state().activeSheetId!;
        render(<Harness />);

        const textarea = document.createElement("textarea");
        document.body.appendChild(textarea);
        dispatchKey("[", {}, textarea);

        expect(state().activeSheetId).toBe(second);
        textarea.remove();
    });
});

describe("useKeymap grid scope", () => {
    beforeEach(() => {
        vi.clearAllMocks();
        document.body.innerHTML = "";
        useFlowStore.setState({ round: null, activeSheetId: null, keymapOverrides: {} });
        freshRound();
        render(<Harness />);
    });

    function focus(tag: "input" | "textarea", className?: string): HTMLElement {
        const el = document.createElement(tag);
        if (className) el.className = className;
        document.body.appendChild(el);
        el.focus();
        return el;
    }

    it("does not format the sheet behind a chrome text field", () => {
        const input = focus("input");
        dispatchKey("b", MOD, input);
        expect(executeCommand).not.toHaveBeenCalled();
    });

    it("formats from the grid's own cell editor", () => {
        const editor = focus("textarea", "handsontableInput");
        dispatchKey("b", MOD, editor);
        expect(executeCommand).toHaveBeenCalledWith("format.toggleBold");
    });

    it("formats when no text field holds focus", () => {
        dispatchKey("b", MOD);
        expect(executeCommand).toHaveBeenCalledWith("format.toggleBold");
    });

    it("still runs an app-scoped chord from a chrome text field", () => {
        const input = focus("input");
        dispatchKey("P", { ...MOD, shiftKey: true }, input);
        expect(executeCommand).toHaveBeenCalledWith("palette.open");
    });
});
