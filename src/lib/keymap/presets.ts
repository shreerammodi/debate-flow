/**
 * Single flat modeless keymap preset.
 *
 * Chord strings are canonical (see resolve.ts / eventToChord). Grid-native
 * gestures (Enter, Alt+Enter, Tab, Esc, arrows) are owned by Handsontable and
 * never appear here; this map holds only app chords.
 */

import type { CommandId } from "@/lib/commands/registry";
import { isMacPlatform } from "@/lib/platform";

import type { Chord, Keymap } from "./types";

/** Platform modifier letter chords: Meta on Mac, Ctrl elsewhere. */
const LETTER_BINDINGS: Record<Chord, CommandId> = (() => {
    const mod = isMacPlatform() ? "Meta" : "Ctrl";
    return {
        // Document chords. Open and New Flow are deliberately absent: Meta+o
        // and Meta+O are both insert commands a debater uses mid-speech, and
        // Meta+n is New Window (below) - flowing and a fresh window both
        // outrank a between-rounds action. The start screen binds bare "o"
        // and "n" instead.
        [`${mod}+n`]: "window.new",
        [`${mod}+s`]: "flow.save",
        [`${mod}+S`]: "flow.saveAs",
        [`${mod}+z`]: "edit.undo",
        [`${mod}+Z`]: "edit.redo",
        [`${mod}+b`]: "format.toggleBold",
        [`${mod}+H`]: "format.toggleHighlight",
        [`${mod}+t`]: "format.toggleCard",
        [`${mod}+g`]: "format.toggleGroup",
        [`${mod}+k`]: "format.toggleKicked",
        [`${mod}+p`]: "sheet.quickSwitch",
        [`${mod}+P`]: "palette.open",
        [`${mod}+A`]: "sheet.newAff",
        [`${mod}+N`]: "sheet.newNeg",
        [`${mod}+r`]: "sheet.rename",
        [`${mod}+,`]: "settings.open",
        [`${mod}+\\`]: "sidebar.toggle",
        [`${mod}+j`]: "rfd.toggle",
        [`${mod}+Backspace`]: "row.delete",
        [`${mod}+o`]: "cell.insert",
        [`${mod}+Alt+o`]: "cell.insertBelow",
        [`${mod}+e`]: "cell.jumpToSource",
        [`${mod}+E`]: "cell.sendToDoc",
        // Bare Meta+m is the Tauri window's native minimize chord, so the move
        // mode takes the shifted one. eventToChord encodes shift in the letter's
        // case, which makes Meta+Shift+m the string "Meta+M".
        [`${mod}+M`]: "cell.move",
        // Default pushes the current row down; rebind to row.insertBelow in
        // Settings to insert underneath instead.
        [`${mod}+O`]: "row.insertAbove",
    };
})();

/** Sheet jumps: Meta+1-9 on Mac, Ctrl+1-9 elsewhere. */
const SHEET_JUMPS: Record<Chord, CommandId> = (() => {
    const mod = isMacPlatform() ? "Meta" : "Ctrl";
    return {
        [`${mod}+1`]: "sheet.jump1",
        [`${mod}+2`]: "sheet.jump2",
        [`${mod}+3`]: "sheet.jump3",
        [`${mod}+4`]: "sheet.jump4",
        [`${mod}+5`]: "sheet.jump5",
        [`${mod}+6`]: "sheet.jump6",
        [`${mod}+7`]: "sheet.jump7",
        [`${mod}+8`]: "sheet.jump8",
        [`${mod}+9`]: "sheet.jump9",
    };
})();

/**
 * Split-view chords. All three live on Alt: the platform modifier + h is the
 * OS's hide-app chord on macOS (see src-tauri/src/menu.rs), so pane focus
 * keeps the vim-style h/l letters under Alt instead.
 */
const SPLIT_BINDINGS: Record<Chord, CommandId> = {
    "Alt+\\": "split.toggle",
    "Alt+h": "split.focusLeft",
    "Alt+l": "split.focusRight",
};

/** The single flat keymap: sheet switching, formatting, and utility chords. */
export const FLAT_KEYMAP: Keymap = {
    name: "default",
    bindings: {
        "]": "sheet.next",
        "[": "sheet.prev",
        "?": "help.open",
        ...LETTER_BINDINGS,
        ...SHEET_JUMPS,
        ...SPLIT_BINDINGS,
    },
};

/**
 * Chords a command carried as a preset default in an earlier version, keyed by
 * command. config.toml records every binding explicitly, so a preset that moves
 * a chord to another command leaves the old value behind in every existing
 * file - and an override outranks a default forever, which kept Meta+N on New
 * flow for every install that predates New window. A chord listed here is a
 * leftover default rather than a choice, so reading the file drops it (see
 * toAppConfig). The cost is that a debater who rebound the command back to its
 * old chord by hand loses that binding once; the alternative is an install that
 * never sees the new default at all.
 */
export const RETIRED_DEFAULTS: Record<string, readonly string[]> = {
    "flow.new": ["Meta+n", "Ctrl+n"],
};

// --- Registry ------------------------------------------------------------------

/** Returns the flat preset keymap. */
export function getPresetKeymap(): Keymap {
    return FLAT_KEYMAP;
}
