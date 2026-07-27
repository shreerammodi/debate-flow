/**
 * InviteChip component tests.
 *
 * Uses the real collab store, reset between tests for isolation. The join path
 * is mocked at the module the toast also calls, so what is asserted is that the
 * chip runs it - not what it does once it is running.
 */

import { render, screen, within } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { act } from "react";
import { describe, it, expect, beforeEach, vi } from "vitest";

vi.mock("@/lib/collab/inbox", () => ({ acceptInvite: vi.fn(async () => {}) }));

import InviteChip from "@/components/collab/InviteChip";
import { acceptInvite } from "@/lib/collab/inbox";
import { useCollabStore } from "@/lib/store/useCollabStore";
import { useFlowStore } from "@/lib/store/useFlowStore";

const ALEX = "a1e0".repeat(16);
const RIN = "b2f0".repeat(16);

const HARVARD = { endpointId: ALEX, roundId: "round_1", label: "Round 3 - Harvard" };
const BRONX = { endpointId: RIN, roundId: "round_2", label: "Round 4 - Bronx" };

beforeEach(() => {
    useCollabStore.getState().reset();
    useFlowStore.setState({
        contacts: {
            [ALEX]: { name: "Alex", role: "partner" },
            [RIN]: { name: "Rin", role: "coach" },
        },
    });
    vi.mocked(acceptInvite).mockClear();
});

describe("InviteChip", () => {
    it("renders nothing at all while nobody has offered a round", () => {
        const { container } = render(<InviteChip />);
        expect(container).toBeEmptyDOMElement();
    });

    it("appears when an invitation lands, so a flow that is open is not a dead end", () => {
        useCollabStore.getState().pushInvite(HARVARD);
        render(<InviteChip />);
        expect(screen.getByTestId("collab-invite-chip")).toHaveTextContent("1 invite");
    });

    it("counts more than one", () => {
        useCollabStore.getState().pushInvite(HARVARD);
        useCollabStore.getState().pushInvite(BRONX);
        render(<InviteChip />);
        expect(screen.getByTestId("collab-invite-chip")).toHaveTextContent("2 invites");
    });

    it("names the partner and the round it offers", async () => {
        useCollabStore.getState().pushInvite(HARVARD);
        useCollabStore.getState().pushInvite(BRONX);
        render(<InviteChip />);

        await userEvent.click(screen.getByTestId("collab-invite-chip"));
        const rows = screen.getAllByTestId("collab-invite-row");
        expect(rows[0]).toHaveTextContent("Alex shared Round 3 - Harvard");
        expect(rows[1]).toHaveTextContent("Rin shared Round 4 - Bronx");
    });

    it("joins through the same path the corner message takes", async () => {
        useCollabStore.getState().pushInvite(HARVARD);
        render(<InviteChip />);

        await userEvent.click(screen.getByTestId("collab-invite-chip"));
        await userEvent.click(screen.getByTestId("collab-invite-join"));
        expect(acceptInvite).toHaveBeenCalledWith(HARVARD);
    });

    it("dismisses exactly the offer that was turned down", async () => {
        useCollabStore.getState().pushInvite(HARVARD);
        useCollabStore.getState().pushInvite(BRONX);
        render(<InviteChip />);

        await userEvent.click(screen.getByTestId("collab-invite-chip"));
        const rows = screen.getAllByTestId("collab-invite-row");
        await userEvent.click(within(rows[0]).getByTestId("collab-invite-dismiss"));

        expect(useCollabStore.getState().invites.map((i) => i.roundId)).toEqual(["round_2"]);
        expect(acceptInvite).not.toHaveBeenCalled();
    });

    it("leaves no trace once the last offer is gone, and reopens collapsed", async () => {
        useCollabStore.getState().pushInvite(HARVARD);
        const { container } = render(<InviteChip />);

        await userEvent.click(screen.getByTestId("collab-invite-chip"));
        await userEvent.click(screen.getByTestId("collab-invite-dismiss"));
        expect(container).toBeEmptyDOMElement();

        act(() => useCollabStore.getState().pushInvite(BRONX));
        expect(screen.queryByTestId("collab-invite-list")).toBeNull();
        expect(screen.getByTestId("collab-invite-chip")).toBeInTheDocument();
    });

    // Keyboard-first is the product value; a mouse-only invitation surface
    // would be the one screen a debater has to reach for the trackpad for.
    it("is reachable and joinable from the keyboard alone", async () => {
        useCollabStore.getState().pushInvite(HARVARD);
        render(<InviteChip />);

        await userEvent.tab();
        expect(screen.getByTestId("collab-invite-chip")).toHaveFocus();

        await userEvent.keyboard("{Enter}");
        expect(screen.getByTestId("collab-invite-list")).toBeInTheDocument();

        // The trigger leads in DOM order, so a Tab walks straight into the panel.
        await userEvent.tab();
        expect(screen.getByTestId("collab-invite-join")).toHaveFocus();

        await userEvent.keyboard("{Enter}");
        expect(acceptInvite).toHaveBeenCalledWith(HARVARD);
    });

    it("takes no focus and opens no dialog of its own", () => {
        useCollabStore.getState().pushInvite(HARVARD);
        render(<InviteChip />);

        expect(document.body).toHaveFocus();
        expect(screen.queryByRole("dialog")).toBeNull();
    });
});
