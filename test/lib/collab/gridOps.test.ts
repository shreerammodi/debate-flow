import { describe, expect, it } from "vitest";

import {
    isReplicatedSource,
    rowOpFromHook,
    textOpsFromChanges,
    type ModelChange,
} from "@/lib/collab/gridOps";
import { modelCol } from "@/lib/grid/colSpace";

describe("isReplicatedSource", () => {
    it("accepts the writes that change text without moving any cell", () => {
        for (const s of [
            "edit",
            "CopyPaste.paste",
            "CopyPaste.cut",
            "UndoRedo.undo",
            "UndoRedo.redo",
            "populateFromArray",
        ]) {
            expect(isReplicatedSource(s)).toBe(true);
        }
    });

    it("refuses a structured write, which names itself at its own call site", () => {
        expect(isReplicatedSource("ebb.structured")).toBe(false);
    });

    it("refuses a source it does not know, and a missing one", () => {
        expect(isReplicatedSource("loadData")).toBe(false);
        expect(isReplicatedSource(undefined)).toBe(false);
        expect(isReplicatedSource(null)).toBe(false);
    });
});

describe("textOpsFromChanges", () => {
    it("turns each changed cell into one independent write", () => {
        const changes: ModelChange[] = [
            [0, modelCol(1), "old", "new"],
            [3, modelCol(0), null, "fresh"],
        ];
        expect(textOpsFromChanges("sheet-1", changes)).toEqual([
            { kind: "cellText", sheetId: "sheet-1", col: 1, row: 0, text: "new" },
            { kind: "cellText", sheetId: "sheet-1", col: 0, row: 3, text: "fresh" },
        ]);
    });

    it("carries an emptied cell as empty text, not as a delete", () => {
        expect(textOpsFromChanges("s", [[2, modelCol(2), "gone", ""]])).toEqual([
            { kind: "cellText", sheetId: "s", col: 2, row: 2, text: "" },
        ]);
    });

    it("drops a change that reports no actual difference", () => {
        expect(textOpsFromChanges("s", [[0, modelCol(0), "same", "same"]])).toEqual([]);
    });

    it("ignores a non-numeric column, which a keyed data source would give", () => {
        expect(textOpsFromChanges("s", [[0, "title", "a", "b"]])).toEqual([]);
    });
});

describe("rowOpFromHook", () => {
    it("expands an amount into one op per row, top down for insert", () => {
        expect(rowOpFromHook("insert", "s", 4, 2, undefined)).toEqual([
            { kind: "insertRow", sheetId: "s", row: 4 },
            { kind: "insertRow", sheetId: "s", row: 5 },
        ]);
    });

    it("removes the same index repeatedly, because each removal closes the gap", () => {
        expect(rowOpFromHook("remove", "s", 4, 3, undefined)).toEqual([
            { kind: "removeRow", sheetId: "s", row: 4 },
            { kind: "removeRow", sheetId: "s", row: 4 },
            { kind: "removeRow", sheetId: "s", row: 4 },
        ]);
    });

    it("accepts the context menu, which has its own source", () => {
        expect(rowOpFromHook("insert", "s", 0, 1, "ContextMenu.rowAbove")).toHaveLength(1);
    });

    it("refuses the spare row Handsontable grows on its own", () => {
        expect(rowOpFromHook("insert", "s", 250, 1, "auto")).toEqual([]);
    });

    it("refuses a nonsense amount rather than looping", () => {
        expect(rowOpFromHook("insert", "s", 0, 0, undefined)).toEqual([]);
        expect(rowOpFromHook("insert", "s", 0, -3, undefined)).toEqual([]);
    });
});
