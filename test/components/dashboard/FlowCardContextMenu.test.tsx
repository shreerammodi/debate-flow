import "fake-indexeddb/auto";
import { render, screen, waitFor } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { describe, it, expect, vi, beforeEach } from "vitest";

vi.mock("sonner", () => ({
    toast: Object.assign(vi.fn(), { success: vi.fn() }),
}));

import FlowCardContextMenu from "@/components/dashboard/FlowCardContextMenu";
import { makeFlowRound, type FlowRound } from "@/lib/model/flow";
import { flowDb } from "@/lib/persistence/flowDb";
import { listFlows, listFlowTrash, persistFlow } from "@/lib/persistence/flowPersistence";

function mk(id: string): FlowRound {
    return { ...makeFlowRound({ role: "aff" }), id, createdAt: 1, updatedAt: 1 };
}

function open(id: string) {
    return userEvent.pointer({
        target: screen.getByTestId(`context-trigger-${id}`),
        keys: "[MouseRight]",
    });
}

beforeEach(async () => {
    await flowDb.flows.clear();
});

describe("FlowCardContextMenu", () => {
    it("opens on right click with the same actions as the kebab menu", async () => {
        render(
            <FlowCardContextMenu id="a" onViewDetails={() => {}} onChanged={() => {}}>
                <div>card</div>
            </FlowCardContextMenu>,
        );
        await open("a");
        expect(await screen.findByTestId("context-details-a")).toBeInTheDocument();
        expect(screen.getByText("Export")).toBeInTheDocument();
        expect(screen.getByTestId("context-delete-a")).toBeInTheDocument();
    });

    it("routes View details to the callback", async () => {
        const onViewDetails = vi.fn();
        render(
            <FlowCardContextMenu id="a" onViewDetails={onViewDetails} onChanged={() => {}}>
                <div>card</div>
            </FlowCardContextMenu>,
        );
        await open("a");
        await userEvent.click(await screen.findByTestId("context-details-a"));
        expect(onViewDetails).toHaveBeenCalledWith("a");
    });

    it("soft-deletes the flow and calls onChanged", async () => {
        await persistFlow(mk("a"));
        const onChanged = vi.fn();
        render(
            <FlowCardContextMenu id="a" onViewDetails={() => {}} onChanged={onChanged}>
                <div>card</div>
            </FlowCardContextMenu>,
        );
        await open("a");
        await userEvent.click(await screen.findByTestId("context-delete-a"));
        await waitFor(() => expect(onChanged).toHaveBeenCalled());
        expect((await listFlows()).length).toBe(0);
        expect((await listFlowTrash()).map((s) => s.id)).toEqual(["a"]);
    });
});
