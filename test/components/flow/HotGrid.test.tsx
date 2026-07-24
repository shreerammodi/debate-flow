import Handsontable from "handsontable/base";
import { registerAllModules } from "handsontable/registry";
import { afterEach, describe, expect, it } from "vitest";

import { applyMeta, collectMeta } from "@/components/flow/HotGrid";
import type { CellMeta, CellSource } from "@/lib/model/flow";

registerAllModules();

const SRC: CellSource = {
    app: "cardmirror",
    token: "cmsrc1abc",
    key: "doc1|perm solves",
    title: "AT - Cap K",
};

/**
 * The two halves of a sheet switch, driven against a real grid: the pane saves
 * the outgoing sheet with collectMeta and loads the incoming one with applyMeta.
 */
describe("collectMeta / applyMeta", () => {
    let hot: Handsontable | null = null;

    afterEach(() => {
        hot?.destroy();
        hot = null;
    });

    function makeHot() {
        const el = document.createElement("div");
        document.body.appendChild(el);
        hot = new Handsontable(el, {
            data: [
                ["a", "b"],
                ["c", "d"],
            ],
            licenseKey: "non-commercial-and-evaluation",
        });
        return hot;
    }

    it("round-trips provenance and decorations through a sheet switch", () => {
        const h = makeHot();
        h.setCellMeta(0, 0, "className", "flow-card");
        h.setCellMeta(0, 0, "source", SRC);
        h.setCellMeta(1, 1, "source", SRC);

        const saved = collectMeta(h);
        expect(saved).toEqual({
            "0,0": { card: true, source: SRC },
            "1,1": { source: SRC },
        });

        // Switch away to a bare sheet, then back.
        applyMeta(h, {}, saved);
        expect(h.getCellMeta(0, 0).className).toBe("");
        expect(h.getCellMeta(0, 0).source).toBeUndefined();
        expect(h.getCellMeta(1, 1).source).toBeUndefined();

        applyMeta(h, saved, {});
        expect(h.getCellMeta(0, 0).className).toBe("flow-card");
        expect(h.getCellMeta(0, 0).source).toEqual(SRC);
        expect(h.getCellMeta(1, 1).source).toEqual(SRC);
        expect(collectMeta(h)).toEqual(saved);
    });

    it("clears provenance a cell keeps in the outgoing sheet but not the incoming one", () => {
        const h = makeHot();
        h.setCellMeta(0, 0, "className", "flow-bold");
        h.setCellMeta(0, 0, "source", SRC);

        const prev: Record<string, CellMeta> = { "0,0": { bold: true, source: SRC } };
        // The same cell is decorated in the incoming sheet, so the prevMeta pass
        // skips it; only the inject pass can strip the stale provenance.
        applyMeta(h, { "0,0": { bold: true } }, prev);

        expect(h.getCellMeta(0, 0).className).toBe("flow-bold");
        expect(h.getCellMeta(0, 0).source).toBeUndefined();
    });

    it("scans the whole grid for orphaned provenance when the outgoing sheet is gone", () => {
        const h = makeHot();
        h.setCellMeta(1, 0, "className", "flow-highlight");
        h.setCellMeta(1, 0, "source", SRC);
        h.setCellMeta(0, 1, "source", SRC);

        applyMeta(h, {}, null);

        expect(h.getCellMeta(1, 0).className).toBe("");
        expect(h.getCellMeta(1, 0).source).toBeUndefined();
        expect(h.getCellMeta(0, 1).source).toBeUndefined();
        expect(collectMeta(h)).toEqual({});
    });
});
