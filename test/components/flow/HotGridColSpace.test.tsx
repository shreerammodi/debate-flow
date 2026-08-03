import { render, waitFor } from "@testing-library/react";
import { afterEach, describe, expect, it } from "vitest";

import HotGrid from "@/components/flow/HotGrid";
import { getActiveHot, notifyGridMutated } from "@/lib/grid/hotInstance";
import { makeFlowRound, makeFlowSheet } from "@/lib/model/flow";
import { useFlowStore } from "@/lib/store/useFlowStore";

// jsdom gives every row zero height, so Handsontable virtualizes nothing and
// each write redraws all MIN_ROWS rows of every column. A sweep therefore
// writes its whole row in one call: setDataAtCell takes the changes as a
// batch, afterChange still converts each column of the batch on its own, and
// the grid is redrawn once instead of once per column.
const SWEEP_MS = 30_000;

/**
 * Every column the grid holds, driven at both ends. A dropped conversion puts
 * a cell one column from where the debater put it, which is silent, reaches
 * the file and reaches a partner.
 */
describe("grid and model columns agree", () => {
    afterEach(() => {
        useFlowStore.setState({
            alignSpeeches: false,
            round: null,
            activeSheetId: null,
            splitSheetId: null,
        });
    });

    async function padded(startSpeechId: string | undefined) {
        const round = makeFlowRound();
        const sheet = { ...makeFlowSheet({ title: "2.", group: "neg", order: 1 }), startSpeechId };
        round.sheets.push(sheet);
        useFlowStore.setState({
            round,
            activeSheetId: sheet.id,
            splitSheetId: null,
            alignSpeeches: true,
        });
        render(<HotGrid sheetId={sheet.id} pane={1} />);
        const hot = await waitFor(() => {
            const h = getActiveHot();
            expect(h).not.toBeNull();
            return h!;
        });
        return { hot, sheetId: sheet.id };
    }

    // undefined is the default neg start (one spacer); the rest walk the pad
    // out to three, which is the widest a Policy round can derive.
    for (const [start, spacers] of [
        [undefined, 1],
        ["2ac", 2],
        ["block", 3],
    ] as const) {
        it(
            `round-trips every cell of a sheet with ${spacers} spacer(s)`,
            async () => {
                const { hot, sheetId } = await padded(start);
                const shown = hot.countCols() - spacers;

                hot.setDataAtCell(
                    Array.from({ length: shown }, (_, c) => [0, spacers + c, `col ${c}`] as const),
                );
                const saved = useFlowStore.getState().round!.sheets.find((s) => s.id === sheetId)!;
                for (let c = 0; c < shown; c++) {
                    expect(saved.data[0][c]).toBe(`col ${c}`);
                }
                expect(saved.data[0].length).toBe(shown);
            },
            SWEEP_MS,
        );

        it(
            `keeps decorations on the cell they were set on with ${spacers} spacer(s)`,
            async () => {
                const { hot, sheetId } = await padded(start);
                const shown = hot.countCols() - spacers;

                hot.setDataAtCell(
                    Array.from({ length: shown }, (_, c) => [0, spacers + c, `col ${c}`] as const),
                );
                for (let c = 0; c < shown; c++) {
                    hot.setCellMeta(0, spacers + c, "className", "flow-bold");
                }
                // A decoration reaches no afterChange hook: the bold command
                // writes the meta and calls this, which is what saves it.
                notifyGridMutated();

                const saved = useFlowStore.getState().round!.sheets.find((s) => s.id === sheetId)!;
                for (let c = 0; c < shown; c++) {
                    expect(saved.meta[`0,${c}`]).toEqual({ bold: true });
                }
                expect(saved.meta["0,-1"]).toBeUndefined();
            },
            SWEEP_MS,
        );
    }
});
