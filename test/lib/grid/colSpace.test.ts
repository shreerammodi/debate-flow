import { describe, expect, it } from "vitest";

import { gridCol, modelCol, toGridCol, toModelCol } from "@/lib/grid/colSpace";

describe("colSpace", () => {
    it("shifts a column by the pane's spacer count in both directions", () => {
        expect(toGridCol(modelCol(0), 1)).toBe(1);
        expect(toModelCol(gridCol(1), 1)).toBe(0);
        expect(toGridCol(modelCol(5), 3)).toBe(8);
        expect(toModelCol(gridCol(8), 3)).toBe(5);
    });

    it("is the identity on an unpadded pane", () => {
        expect(toGridCol(modelCol(4), 0)).toBe(4);
        expect(toModelCol(gridCol(4), 0)).toBe(4);
    });

    it("round-trips every column of a padded pane", () => {
        for (const spacers of [0, 1, 2, 7]) {
            for (let c = 0; c < 12; c++) {
                expect(toModelCol(toGridCol(modelCol(c), spacers), spacers)).toBe(c);
            }
        }
    });

    it("floors a spacer at the sheet's first column rather than going negative", () => {
        // A grid column inside the pad belongs to no cell of the sheet. Callers
        // clamp before converting; a negative model column would index a row
        // from its end, which is a cell the debater never pointed at.
        expect(toModelCol(gridCol(0), 2)).toBe(0);
        expect(toModelCol(gridCol(1), 2)).toBe(0);
    });
});
