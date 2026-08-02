import Handsontable from "handsontable/base";
import { registerAllModules } from "handsontable/registry";
import { afterEach, beforeEach, describe, expect, it } from "vitest";

import { insertCell } from "@/lib/grid/cellShift";
import { gridCol } from "@/lib/grid/colSpace";
import {
    attachMetaUndo,
    onRedoStackChange,
    onUndoStackChange,
    resetMetaUndo,
    restoreMetaRedo,
    restoreMetaUndo,
    snapshotClasses,
} from "@/lib/grid/metaUndo";
import type { CellSource } from "@/lib/model/flow";

const SRC: CellSource = { app: "cardmirror", token: "cmsrc1abc", key: "doc1|perm solves" };

registerAllModules();

/**
 * Handsontable's undo stack records setDataAtCell and ignores setCellMeta, so
 * the module rides the documented undo hooks on a real grid rather than a stub.
 */
describe("metaUndo on a live grid", () => {
    let hot: Handsontable | null = null;

    beforeEach(() => {
        resetMetaUndo();
    });
    afterEach(() => {
        hot?.destroy();
        hot = null;
        resetMetaUndo();
    });

    function makeHot() {
        const el = document.createElement("div");
        document.body.appendChild(el);
        hot = new Handsontable(el, {
            data: [["a"], ["b"], ["c"]],
            undo: true,
            licenseKey: "non-commercial-and-evaluation",
            afterUndoStackChange: onUndoStackChange,
            afterRedoStackChange: onRedoStackChange,
            afterUndo: () => restoreMetaUndo(hot!),
            afterRedo: () => restoreMetaRedo(hot!),
        });
        return hot;
    }

    /** What commands.runInsertCell does, minus the store plumbing. */
    function doInsert(h: Handsontable, row: number, col: number) {
        const before = snapshotClasses(h, [col]);
        h.setDataAtCell(insertCell(h, row, gridCol(col)));
        attachMetaUndo({ cols: [col], before, after: snapshotClasses(h, [col]) });
    }

    it("returns a decoration to its cell when the insert that moved it is undone", () => {
        const h = makeHot();
        h.setCellMeta(1, 0, "className", "flow-bold");

        doInsert(h, 1, 0);
        expect(h.getDataAtCell(2, 0)).toBe("b");
        expect(h.getCellMeta(2, 0).className).toBe("flow-bold");

        h.getPlugin("undoRedo").undo();

        expect(h.getDataAtCell(1, 0)).toBe("b");
        expect(h.getCellMeta(1, 0).className).toBe("flow-bold");
        expect(h.getCellMeta(2, 0).className).toBe("");
    });

    it("re-applies the decoration on redo", () => {
        const h = makeHot();
        h.setCellMeta(1, 0, "className", "flow-bold");

        doInsert(h, 1, 0);
        h.getPlugin("undoRedo").undo();
        h.getPlugin("undoRedo").redo();

        expect(h.getDataAtCell(2, 0)).toBe("b");
        expect(h.getCellMeta(2, 0).className).toBe("flow-bold");
        expect(h.getCellMeta(1, 0).className).toBe("");
    });

    it("leaves an unattached action alone", () => {
        const h = makeHot();
        h.setCellMeta(0, 0, "className", "flow-bold");

        h.setDataAtCell(0, 0, "z");
        h.getPlugin("undoRedo").undo();

        expect(h.getDataAtCell(0, 0)).toBe("a");
        expect(h.getCellMeta(0, 0).className).toBe("flow-bold");
    });

    it("returns provenance to its cell when the insert that moved it is undone", () => {
        const h = makeHot();
        h.setCellMeta(1, 0, "source", SRC);

        doInsert(h, 1, 0);
        expect(h.getCellMeta(2, 0).source).toEqual(SRC);
        expect(h.getCellMeta(1, 0).source).toBeUndefined();

        h.getPlugin("undoRedo").undo();

        expect(h.getDataAtCell(1, 0)).toBe("b");
        expect(h.getCellMeta(1, 0).source).toEqual(SRC);
        expect(h.getCellMeta(2, 0).source).toBeUndefined();
    });

    it("re-applies provenance on redo", () => {
        const h = makeHot();
        h.setCellMeta(1, 0, "className", "flow-card");
        h.setCellMeta(1, 0, "source", SRC);

        doInsert(h, 1, 0);
        h.getPlugin("undoRedo").undo();
        h.getPlugin("undoRedo").redo();

        expect(h.getDataAtCell(2, 0)).toBe("b");
        expect(h.getCellMeta(2, 0).className).toBe("flow-card");
        expect(h.getCellMeta(2, 0).source).toEqual(SRC);
        expect(h.getCellMeta(1, 0).className).toBe("");
        expect(h.getCellMeta(1, 0).source).toBeUndefined();
    });
});

describe("snapshotClasses", () => {
    it("records only the decorated cells of the named columns", () => {
        const grid = {
            countRows: () => 3,
            countCols: () => 2,
            getDataAtCell: () => null,
            getCellMeta: (r: number, c: number) =>
                r === 1 && c === 0 ? { className: "flow-bold" } : { className: "" },
            setCellMeta: () => {},
        };

        expect(snapshotClasses(grid, [0, 1])).toEqual([[1, 0, "flow-bold"]]);
    });

    it("records a sourced cell even when it carries no decoration", () => {
        const grid = {
            countRows: () => 2,
            countCols: () => 1,
            getDataAtCell: () => null,
            getCellMeta: (r: number) => (r === 1 ? { source: SRC } : {}),
            setCellMeta: () => {},
        };

        expect(snapshotClasses(grid, [0])).toEqual([[1, 0, "", SRC]]);
    });
});
