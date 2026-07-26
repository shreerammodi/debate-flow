import "fake-indexeddb/auto";
import { render, screen, waitFor } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { describe, it, expect, vi, beforeEach } from "vitest";

const push = vi.fn();
vi.mock("next/navigation", () => ({ useRouter: () => ({ push }) }));

import NewFlowButton from "@/components/dashboard/NewFlowButton";
import { flowDb } from "@/lib/persistence/flowDb";

beforeEach(async () => {
    push.mockReset();
    await flowDb.flows.clear();
});

describe("NewFlowButton", () => {
    it("creates a policy flow and navigates to it", async () => {
        render(<NewFlowButton />);
        await userEvent.click(screen.getByTestId("new-flow"));
        await userEvent.click(await screen.findByTestId("new-flow-policy"));
        await waitFor(() => expect(push).toHaveBeenCalledTimes(1));
        const arg = push.mock.calls[0][0] as string;
        expect(arg).toMatch(/^\/flow\?id=round_/);
        const rounds = await flowDb.flows.toArray();
        expect(rounds).toHaveLength(1);
        expect(rounds[0].event).toBe("policy");
        expect(rounds[0].firstSide).toBe("aff");
    });

    it("creates a pf round with the chosen speaking order", async () => {
        render(<NewFlowButton />);
        await userEvent.click(screen.getByTestId("new-flow"));
        await userEvent.click(await screen.findByTestId("new-flow-pf"));
        await userEvent.click(await screen.findByTestId("new-flow-pf-neg"));
        await waitFor(() => expect(push).toHaveBeenCalledTimes(1));
        const rounds = await flowDb.flows.toArray();
        expect(rounds).toHaveLength(1);
        expect(rounds[0].event).toBe("pf");
        expect(rounds[0].firstSide).toBe("neg");
    });

    it("opens a neg-first pf round on the neg sheet", async () => {
        render(<NewFlowButton />);
        await userEvent.click(screen.getByTestId("new-flow"));
        await userEvent.click(await screen.findByTestId("new-flow-pf"));
        await userEvent.click(await screen.findByTestId("new-flow-pf-neg"));
        await waitFor(() => expect(push).toHaveBeenCalledTimes(1));
        const rounds = await flowDb.flows.toArray();
        const flowSheets = rounds[0].sheets.filter((s) => s.kind !== "cx");
        expect(flowSheets.map((s) => s.group)).toEqual(["neg"]);
    });

    it("creates an ld round with aff-first order", async () => {
        render(<NewFlowButton />);
        await userEvent.click(screen.getByTestId("new-flow"));
        await userEvent.click(await screen.findByTestId("new-flow-ld"));
        await waitFor(() => expect(push).toHaveBeenCalledTimes(1));
        const rounds = await flowDb.flows.toArray();
        expect(rounds).toHaveLength(1);
        expect(rounds[0].event).toBe("ld");
        expect(rounds[0].firstSide).toBe("aff");
    });
});
