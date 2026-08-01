/**
 * SessionChip component tests.
 *
 * Uses the real collab store, reset between tests for isolation.
 */

import { render, screen, within } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { act } from "react";
import { describe, it, expect, beforeEach, vi } from "vitest";

import SessionChip from "@/components/collab/SessionChip";
import { type CollabPeerView, useCollabStore } from "@/lib/store/useCollabStore";

const { disconnectPeer, endSession } = vi.hoisted(() => ({
    disconnectPeer: vi.fn(async () => {}),
    endSession: vi.fn(async () => {}),
}));

vi.mock("@/lib/collab/runtime", () => ({ disconnectPeer, endSession }));

const ALEX: CollabPeerView = {
    endpointId: "alex-endpoint",
    name: "Alex",
    role: "partner",
    connectionType: "direct",
};

const RIN: CollabPeerView = {
    endpointId: "rin-endpoint",
    name: "Rin",
    role: "coach",
    connectionType: "relayed",
};

/** Puts the store in a live session with the given peers. */
function live(peers: CollabPeerView[] = [ALEX, RIN]) {
    useCollabStore.setState({ status: "connected", peers });
}

beforeEach(() => {
    useCollabStore.getState().reset();
    disconnectPeer.mockClear();
    endSession.mockClear();
});

describe("SessionChip", () => {
    it("renders nothing at all while shared editing is off", () => {
        const { container } = render(<SessionChip />);
        expect(container).toBeEmptyDOMElement();
        expect(screen.queryByTestId("collab-chip")).toBeNull();
    });

    it("leaves no trace once a live session ends, and reopens collapsed", async () => {
        const user = userEvent.setup();
        live();
        const { container } = render(<SessionChip />);
        await user.click(screen.getByTestId("collab-chip"));
        expect(screen.getByTestId("collab-chip-peers")).toBeInTheDocument();

        act(() => useCollabStore.getState().reset());
        expect(container).toBeEmptyDOMElement();

        act(() => useCollabStore.setState({ status: "connecting" }));
        expect(screen.getByTestId("collab-chip")).toBeInTheDocument();
        expect(screen.queryByTestId("collab-chip-peers")).toBeNull();
    });

    it("shows the connection state and the peer count while collapsed", () => {
        live();
        render(<SessionChip />);

        const chip = screen.getByTestId("collab-chip");
        expect(chip).toHaveTextContent("Connected");
        expect(chip).toHaveTextContent("2 peers");
        expect(screen.queryByTestId("collab-chip-peers")).toBeNull();
    });

    it("names a single peer in the singular", () => {
        live([ALEX]);
        render(<SessionChip />);
        expect(screen.getByTestId("collab-chip")).toHaveTextContent("1 peer");
    });

    it("reports a session that is still dialing", () => {
        useCollabStore.setState({ status: "connecting", peers: [] });
        render(<SessionChip />);
        expect(screen.getByTestId("collab-chip")).toHaveTextContent("Connecting");
    });

    it("reports a session that dropped and is retrying", () => {
        useCollabStore.setState({ status: "reconnecting", peers: [ALEX] });
        render(<SessionChip />);
        expect(screen.getByTestId("collab-chip")).toHaveTextContent("Reconnecting");
    });

    it("expands on click into one row per peer, and collapses again", async () => {
        const user = userEvent.setup();
        live();
        render(<SessionChip />);

        await user.click(screen.getByTestId("collab-chip"));
        expect(screen.getAllByTestId("collab-peer-row")).toHaveLength(2);

        await user.click(screen.getByTestId("collab-chip"));
        expect(screen.queryByTestId("collab-chip-peers")).toBeNull();
    });

    it("gives a partner the 'edit' badge and a coach 'view'", async () => {
        const user = userEvent.setup();
        live();
        render(<SessionChip />);
        await user.click(screen.getByTestId("collab-chip"));

        const [partner, coach] = screen.getAllByTestId("collab-peer-row");
        expect(within(partner).getByText("Alex")).toBeInTheDocument();
        expect(within(partner).getByTestId("collab-peer-role")).toHaveTextContent("edit");
        expect(within(coach).getByText("Rin")).toBeInTheDocument();
        expect(within(coach).getByTestId("collab-peer-role")).toHaveTextContent("view");
    });

    it("tells a coach their own side is view only", async () => {
        const user = userEvent.setup();
        live();
        useCollabStore.getState().setSelfRole("coach");
        render(<SessionChip />);
        await user.click(screen.getByTestId("collab-chip"));

        expect(screen.getByTestId("collab-self-role")).toHaveTextContent("viewing this round");
    });

    it("says nothing about a side that can edit, which needs no explaining", async () => {
        const user = userEvent.setup();
        live();
        render(<SessionChip />);
        await user.click(screen.getByTestId("collab-chip"));

        expect(screen.queryByTestId("collab-self-role")).toBeNull();
    });

    it("shows the connection type per peer, so a relayed link is disclosed", async () => {
        const user = userEvent.setup();
        live();
        render(<SessionChip />);
        await user.click(screen.getByTestId("collab-chip"));

        const [alex, rin] = screen.getAllByTestId("collab-peer-row");
        expect(within(alex).getByTestId("collab-peer-connection")).toHaveTextContent("direct");
        expect(within(rin).getByTestId("collab-peer-connection")).toHaveTextContent("relayed");
    });

    it("disconnects one peer by endpoint id", async () => {
        const user = userEvent.setup();
        live();
        render(<SessionChip />);

        await user.click(screen.getByTestId("collab-chip"));
        const [, rin] = screen.getAllByTestId("collab-peer-row");
        await user.click(within(rin).getByTestId("collab-peer-disconnect"));

        expect(disconnectPeer).toHaveBeenCalledWith(RIN.endpointId);
    });

    it("ends the session", async () => {
        const user = userEvent.setup();
        live();
        render(<SessionChip />);

        await user.click(screen.getByTestId("collab-chip"));
        await user.click(screen.getByTestId("collab-end-session"));

        expect(endSession).toHaveBeenCalledTimes(1);
    });

    describe("never interrupting the grid", () => {
        it("takes no focus when it appears", () => {
            live();
            render(<SessionChip />);
            expect(document.activeElement).toBe(document.body);
        });

        it("is not a modal", async () => {
            const user = userEvent.setup();
            live();
            render(<SessionChip />);
            await user.click(screen.getByTestId("collab-chip"));

            expect(screen.queryByRole("dialog")).toBeNull();
            expect(screen.getByTestId("collab-chip-peers")).not.toHaveAttribute("aria-modal");
        });

        it("autofocuses nothing inside the expanded panel", async () => {
            const user = userEvent.setup();
            live();
            render(<SessionChip />);

            const chip = screen.getByTestId("collab-chip");
            await user.click(chip);

            // Focus stays on the control the user pressed.
            expect(document.activeElement).toBe(chip);
            expect(screen.getByTestId("collab-chip-peers")).not.toContainElement(
                document.activeElement as HTMLElement,
            );
        });

        it("does not trap focus: Tab walks out of the expanded panel", async () => {
            const user = userEvent.setup();
            live();
            render(
                <>
                    <SessionChip />
                    <button type="button" data-testid="outside">
                        Grid
                    </button>
                </>,
            );

            const chip = screen.getByTestId("collab-chip");
            await user.click(chip);
            const panel = screen.getByTestId("collab-chip-peers");

            // Walk forward off the chip and through its controls; a trap would
            // cycle back into the panel instead of reaching the grid.
            for (let i = 0; i < 8 && chip.parentElement!.contains(document.activeElement); i++) {
                await user.tab();
            }

            expect(panel).toBeInTheDocument();
            expect(document.activeElement).toBe(screen.getByTestId("outside"));
        });
    });

    it("is reachable by keyboard: the collapsed chip is a real button", async () => {
        const user = userEvent.setup();
        live();
        render(<SessionChip />);

        const chip = screen.getByTestId("collab-chip");
        expect(chip.tagName).toBe("BUTTON");
        expect(chip).toHaveAttribute("aria-expanded", "false");

        await user.tab();
        expect(document.activeElement).toBe(chip);

        await user.keyboard("{Enter}");
        expect(screen.getByTestId("collab-chip-peers")).toBeInTheDocument();
        expect(chip).toHaveAttribute("aria-expanded", "true");
    });
});
