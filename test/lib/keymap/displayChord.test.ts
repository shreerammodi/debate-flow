import { describe, it, expect, beforeEach } from "vitest";

import { prettyChord, buildChordMap, keyHintFor } from "@/lib/keymap/displayChord";
import { effectiveKeymap } from "@/lib/keymap/useKeymap";
import { isMacPlatform } from "@/lib/platform";
import { useFlowStore } from "@/lib/store/useFlowStore";

describe("displayChord", () => {
    beforeEach(() => {
        useFlowStore.setState({ keymapOverrides: {} });
    });

    it("prettifies modifier chords", () => {
        expect(prettyChord("Meta+Shift+ArrowUp")).toBe("Cmd-Shift-Up");
        expect(prettyChord("Escape")).toBe("Esc");
    });

    it("renders a shift-bearing uppercase letter chord with Shift", () => {
        // An uppercase single letter encodes Shift (eventToChord rule), so the
        // hint must spell it out rather than show a bare "Cmd-X".
        expect(prettyChord("Meta+X")).toBe("Cmd-Shift-X");
        expect(prettyChord("Meta+Z")).toBe("Cmd-Shift-Z");
        // Lowercase letters carry no Shift and render as-is.
        expect(prettyChord("Meta+z")).toBe("Cmd-z");
    });

    it("maps a bound command to its chord", () => {
        const map = buildChordMap();
        expect(map["sheet.next"]).toBe(`${isMacPlatform() ? "Meta" : "Ctrl"}+]`);
    });

    it("returns a pretty hint for a bound command", () => {
        expect(keyHintFor("sheet.next")).toBe(prettyChord(buildChordMap()["sheet.next"]!));
    });

    it("reflects a user override", () => {
        useFlowStore.setState({
            keymapOverrides: { "sheet.next": "Meta+J" },
        });
        expect(keyHintFor("sheet.next")).toBe("Cmd-Shift-J");
    });

    // One chord per command is structural, not a tie `buildChordMap` breaks:
    // an override is keyed by command, so binding a command replaces its chord
    // rather than adding one. This pins that, because the map's "keep the first"
    // arm reads like an ordering contract and there is no ordering to have.
    it("gives a command one chord, because an override replaces rather than adds", () => {
        useFlowStore.setState({ keymapOverrides: { "sheet.next": "Meta+J" } });

        const chords = Object.entries(effectiveKeymap().bindings)
            .filter(([, cmd]) => cmd === "sheet.next")
            .map(([chord]) => chord);

        expect(chords).toEqual(["Meta+J"]);
        expect(buildChordMap()["sheet.next"]).toBe("Meta+J");
    });
});
