/**
 * SessionControls component tests.
 *
 * The buttons are the visible route to a session, so what each one dispatches
 * - and when it refuses to - is the whole contract.
 */

import { render, screen } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { act } from "react";
import { beforeEach, describe, expect, it, vi } from "vitest";

import SessionControls from "@/components/collab/SessionControls";
import type { FlowRound } from "@/lib/model/flow";
import { useCollabStore } from "@/lib/store/useCollabStore";
import { useFlowStore } from "@/lib/store/useFlowStore";

const executeCommand = vi.hoisted(() => vi.fn());

vi.mock("@/lib/commands/commands", () => ({ executeCommand }));

const ROUND = { id: "r1" } as FlowRound;

beforeEach(() => {
    executeCommand.mockClear();
    useFlowStore.setState({ round: null, contacts: {} });
    useCollabStore.getState().reset();
});

describe("SessionControls", () => {
    it("shares the open round and copies the invite", async () => {
        useFlowStore.setState({ round: ROUND });
        render(<SessionControls />);

        await userEvent.click(screen.getByTestId("session-share"));
        expect(executeCommand).toHaveBeenCalledWith("collab.share");
    });

    it("refuses to share with no flow open", () => {
        render(<SessionControls />);
        expect(screen.getByTestId("session-share")).toBeDisabled();
        expect(screen.getByTestId("session-invite")).toBeDisabled();
    });

    it("offers a saved partner only once there is one", () => {
        useFlowStore.setState({ round: ROUND });
        const { rerender } = render(<SessionControls />);
        expect(screen.getByTestId("session-invite")).toBeDisabled();

        act(() => {
            useFlowStore.setState({ contacts: { alex: { name: "Alex", role: "partner" } } });
        });
        rerender(<SessionControls />);
        expect(screen.getByTestId("session-invite")).not.toBeDisabled();
    });

    it("joins from the clipboard with no flow open", async () => {
        render(<SessionControls />);

        await userEvent.click(screen.getByTestId("session-join"));
        expect(executeCommand).toHaveBeenCalledWith("collab.join");
    });

    it("offers to end only while a session is running", async () => {
        render(<SessionControls />);
        expect(screen.queryByTestId("session-end")).toBeNull();

        act(() => useCollabStore.getState().setStatus("connected"));
        await userEvent.click(screen.getByTestId("session-end"));
        expect(executeCommand).toHaveBeenCalledWith("collab.end");
    });
});
