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

    it("round-trips every cell of a padded pane, both directions", () => {
        for (const spacers of [0, 1, 2, 7]) {
            for (let c = 0; c < 12; c++) {
                expect(toModelCol(toGridCol(modelCol(c), spacers), spacers)).toBe(c);
            }
            for (let c = spacers; c < 12; c++) {
                const model = toModelCol(gridCol(c), spacers);
                expect(model).not.toBeNull();
                expect(toGridCol(model!, spacers)).toBe(c);
            }
        }
    });

    it("gives a column inside the pad no cell at all", () => {
        // Clamping would answer 0, which is a real addressable cell, so a
        // caller that forgot to exclude the pad would act on the sheet's first
        // column instead of failing.
        expect(toModelCol(gridCol(0), 2)).toBeNull();
        expect(toModelCol(gridCol(1), 2)).toBeNull();
        expect(toModelCol(gridCol(2), 2)).toBe(0);
    });

    it("keeps the two spaces apart at the type level", () => {
        // The whole deliverable is this boundary, and every runtime assertion
        // above passes just as well if the brands dissolve into plain numbers.
        // @ts-expect-error a grid column is not a cell index
        toGridCol(gridCol(1), 0);
        // @ts-expect-error a cell index is not a grid column
        toModelCol(modelCol(1), 0);
        // @ts-expect-error a bare number is neither until it is named
        toGridCol(1, 0);
        // @ts-expect-error relabelling across the spaces would drop the shift
        modelCol(gridCol(1));
        // @ts-expect-error the same, in the other direction
        gridCol(modelCol(1));
        expect(true).toBe(true);
    });
});
