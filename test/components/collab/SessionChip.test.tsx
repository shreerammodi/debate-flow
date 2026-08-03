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
    role: "editor",
    connectionType: "direct",
};

const RIN: CollabPeerView = {
    endpointId: "rin-endpoint",
    name: "Rin",
    role: "viewer",
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

    it("counts the partners who are here while collapsed", () => {
        live();
        render(<SessionChip />);

        const chip = screen.getByTestId("collab-chip");
        expect(chip).toHaveTextContent("Connected to 2 partners");
        expect(screen.queryByTestId("collab-chip-peers")).toBeNull();
    });

    it("names a single partner rather than counting them", () => {
        live([ALEX]);
        render(<SessionChip />);
        expect(screen.getByTestId("collab-chip")).toHaveTextContent("Connected to Alex");
    });

    it("notes a relayed link quietly, on the same line", () => {
        live([RIN]);
        render(<SessionChip />);
        expect(screen.getByTestId("collab-chip")).toHaveTextContent("Connected to Rin, relayed");
    });

    it("says a session nobody has joined is open, not broken", () => {
        useCollabStore.setState({ status: "connecting", peers: [], pending: [] });
        render(<SessionChip />);
        expect(screen.getByTestId("collab-chip")).toHaveTextContent("Waiting to be joined");
    });

    it("names the partner it is waiting for", () => {
        useCollabStore.setState({
            status: "connecting",
            peers: [],
            pending: [{ endpointId: "sam-endpoint", name: "Sam", unreachable: false }],
        });
        render(<SessionChip />);
        expect(screen.getByTestId("collab-chip")).toHaveTextContent("Waiting for Sam");
    });

    it("says so when it cannot reach that partner", () => {
        useCollabStore.setState({
            status: "reconnecting",
            peers: [],
            pending: [{ endpointId: "sam-endpoint", name: "Sam", unreachable: true }],
        });
        render(<SessionChip />);
        expect(screen.getByTestId("collab-chip")).toHaveTextContent("Can't reach Sam");
    });

    it("spells out what to do about a partner it cannot reach", async () => {
        const user = userEvent.setup();
        useCollabStore.setState({
            status: "connected",
            peers: [ALEX],
            pending: [{ endpointId: "sam-endpoint", name: "Sam", unreachable: true }],
        });
        render(<SessionChip />);

        await user.click(screen.getByTestId("collab-chip"));
        expect(screen.getAllByTestId("collab-peer-row")).toHaveLength(1);
        expect(screen.getByTestId("collab-pending-row")).toHaveTextContent(
            "Can't reach Sam. You both need internet, or the same wifi.",
        );
    });

    it("spells out that a partner has not opened the round yet", async () => {
        const user = userEvent.setup();
        useCollabStore.setState({
            status: "connecting",
            peers: [],
            pending: [{ endpointId: "sam-endpoint", name: "Sam", unreachable: false }],
        });
        render(<SessionChip />);

        await user.click(screen.getByTestId("collab-chip"));
        expect(screen.getByTestId("collab-pending-row")).toHaveTextContent(
            "Waiting for Sam to open this round",
        );
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

    it("badges an editor 'edit' and a viewer 'view'", async () => {
        const user = userEvent.setup();
        live();
        render(<SessionChip />);
        await user.click(screen.getByTestId("collab-chip"));

        const [editor, viewer] = screen.getAllByTestId("collab-peer-row");
        expect(within(editor).getByText("Alex")).toBeInTheDocument();
        expect(within(editor).getByTestId("collab-peer-role")).toHaveTextContent("edit");
        expect(within(viewer).getByText("Rin")).toBeInTheDocument();
        expect(within(viewer).getByTestId("collab-peer-role")).toHaveTextContent("view");
    });

    it("tells a viewer their own side is view only", async () => {
        const user = userEvent.setup();
        live();
        useCollabStore.getState().setSelfRole("viewer");
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
