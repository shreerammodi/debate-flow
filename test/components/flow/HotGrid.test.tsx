import { render, screen } from "@testing-library/react";
import Handsontable from "handsontable/base";
import { registerAllModules } from "handsontable/registry";
import { afterEach, describe, expect, it } from "vitest";

import HotGrid, { applyMeta, collectMeta, COL_WIDTH } from "@/components/flow/HotGrid";
import { makeFlowRound, makeFlowSheet, type CellMeta, type CellSource } from "@/lib/model/flow";
import { useFlowStore } from "@/lib/store/useFlowStore";

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

/**
 * Aligning pads the pane by the speeches a sheet does not show, so the same
 * speech lands at the same place on every sheet of the round.
 */
describe("speech alignment", () => {
    const round = makeFlowRound();
    // makeFlowRound opens a policy round with the cx sheet first and one aff
    // flow sheet after it; the round needs a neg sheet to have an offset one.
    const affSheet = round.sheets[1];
    const negSheet = makeFlowSheet({ title: "2.", group: "neg", order: 1 });
    round.sheets.push(negSheet);

    afterEach(() => {
        useFlowStore.setState({ alignSpeeches: false });
    });

    function padOf(sheetId: string, alignSpeeches: boolean): string {
        useFlowStore.setState({
            round,
            activeSheetId: sheetId,
            splitSheetId: null,
            alignSpeeches,
        });
        render(<HotGrid sheetId={sheetId} pane={1} />);
        return screen.getByTestId("grid-pad").style.paddingLeft;
    }

    it("pads a neg sheet by the speech that opens the round", () => {
        expect(padOf(negSheet.id, true)).toBe(`${COL_WIDTH}px`);
    });

    it("leaves the sheet that opens the round flush", () => {
        expect(padOf(affSheet.id, true)).toBe("0px");
    });

    it("pads nothing while the setting is off", () => {
        expect(padOf(negSheet.id, false)).toBe("0px");
    });
});
