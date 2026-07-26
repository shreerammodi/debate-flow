import { render, screen } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { describe, it, expect, vi, beforeEach } from "vitest";

import ExportMenu from "@/components/flow/ExportMenu";
import { TooltipProvider } from "@/components/ui/tooltip";
import { makeFlowRound } from "@/lib/model/flow";
import { useFlowStore } from "@/lib/store/useFlowStore";

vi.mock("@/lib/export/xlsx", () => ({
    downloadXlsx: vi.fn().mockResolvedValue(undefined),
}));

beforeEach(() => {
    useFlowStore.getState().loadRound(makeFlowRound({}));
});

describe("ExportMenu", () => {
    it("opens on click and offers Excel", async () => {
        const user = userEvent.setup();
        render(
            <TooltipProvider>
                <ExportMenu />
            </TooltipProvider>,
        );
        await user.click(screen.getByTestId("export-btn"));
        expect(await screen.findByTestId("export-excel")).toBeInTheDocument();
    });

    it("no longer offers JSON, because a .ebb file already is the round's JSON", async () => {
        const user = userEvent.setup();
        render(
            <TooltipProvider>
                <ExportMenu />
            </TooltipProvider>,
        );
        await user.click(screen.getByTestId("export-btn"));
        await screen.findByTestId("export-excel");
        expect(screen.queryByTestId("export-json")).not.toBeInTheDocument();
    });

    it("Excel item invokes downloadXlsx", async () => {
        const user = userEvent.setup();
        const { downloadXlsx } = await import("@/lib/export/xlsx");
        render(
            <TooltipProvider>
                <ExportMenu />
            </TooltipProvider>,
        );
        await user.click(screen.getByTestId("export-btn"));
        await user.click(await screen.findByTestId("export-excel"));
        expect(downloadXlsx).toHaveBeenCalled();
    });
});
