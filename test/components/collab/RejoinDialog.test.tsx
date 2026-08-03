/**
 * RejoinDialog component tests.
 *
 * The dialog is the whole route the one question a join asks travels, so what
 * reaches the screen and what the join promise settles to are the assertions.
 */

import { render, screen } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { act } from "react";
import { beforeEach, describe, expect, it } from "vitest";

import RejoinDialog from "@/components/collab/RejoinDialog";
import { useFlowStore } from "@/lib/store/useFlowStore";
import { askToRejoin, useRejoinDialog } from "@/lib/store/useRejoinDialog";

/** What iroh hands back, and what a contact is saved under. */
const ALEX = "a".repeat(64);

beforeEach(() => {
    useRejoinDialog.setState({ open: false, ask: null });
    useFlowStore.setState({ contacts: { [ALEX]: { name: "Alex" } } });
});

describe("the confirmation a join asks for a round already here", () => {
    function ask(round = "round-3"): Promise<boolean> {
        let answer!: Promise<boolean>;
        act(() => {
            answer = askToRejoin({ round, endpointId: ALEX });
        });
        return answer;
    }

    it("stays out of the way until a join has something to ask", () => {
        render(<RejoinDialog />);
        expect(screen.queryByTestId("rejoin-dialog")).toBeNull();
    });

    it("names the round and the peer asking for a place in it", () => {
        render(<RejoinDialog />);
        void ask("Berkeley B vs Harvard D");

        expect(screen.getByTestId("rejoin-dialog").textContent).toContain(
            "Berkeley B vs Harvard D",
        );
        expect(screen.getByTestId("rejoin-add").textContent).toBe("Add Alex");
    });

    it("grants the peer only on the answer that says so", async () => {
        render(<RejoinDialog />);
        const answer = ask();

        await userEvent.click(screen.getByTestId("rejoin-add"));
        expect(await answer).toBe(true);
    });

    it("declines on the cancel", async () => {
        render(<RejoinDialog />);
        const answer = ask();

        await userEvent.click(screen.getByTestId("rejoin-cancel"));
        expect(await answer).toBe(false);
    });

    // The whole question exists for a debater who does not recognise the round,
    // so the answer their reflexes give has to be the one that adds nobody.
    it("holds the focus on the cancel, and takes Escape as one", async () => {
        render(<RejoinDialog />);
        const answer = ask();

        expect(screen.getByTestId("rejoin-cancel")).toHaveFocus();
        await userEvent.keyboard("{Escape}");
        expect(await answer).toBe(false);
    });

    it("declines when it leaves the tree, so a join is never left waiting", async () => {
        const { unmount } = render(<RejoinDialog />);
        const answer = ask();

        unmount();
        expect(await answer).toBe(false);
    });

    it("settles the first question when a second one is asked over it", async () => {
        render(<RejoinDialog />);
        const first = ask("round-3");
        const second = ask("round-4");

        await userEvent.click(screen.getByTestId("rejoin-add"));
        expect(await first).toBe(false);
        expect(await second).toBe(true);
    });
});
