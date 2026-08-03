import { render, screen } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { act } from "react";
import { beforeEach, describe, expect, it } from "vitest";

import JoinDialog from "@/components/collab/JoinDialog";
import { askForCode, useJoinDialog } from "@/lib/store/useJoinDialog";

function ask(): Promise<string | null> {
    let answer!: Promise<string | null>;
    act(() => {
        answer = askForCode();
    });
    return answer;
}

beforeEach(() => {
    useJoinDialog.getState().close();
});

describe("JoinDialog", () => {
    it("renders nothing until a join asks", () => {
        render(<JoinDialog />);
        expect(screen.queryByTestId("join-code-field")).toBeNull();
    });

    it("takes a code typed without the dash", async () => {
        render(<JoinDialog />);
        const answer = ask();
        await userEvent.type(screen.getByTestId("join-code-field"), "k7qm3xpv");
        await userEvent.click(screen.getByTestId("join-code-submit"));
        expect(await answer).toBe("K7QM3XPV");
    });

    it("takes a code typed with the dash", async () => {
        render(<JoinDialog />);
        const answer = ask();
        await userEvent.type(screen.getByTestId("join-code-field"), "K7QM-3XPV");
        await userEvent.click(screen.getByTestId("join-code-submit"));
        expect(await answer).toBe("K7QM3XPV");
    });

    it("joins on Enter, because the field holds one value", async () => {
        render(<JoinDialog />);
        const answer = ask();
        await userEvent.type(screen.getByTestId("join-code-field"), "K7QM3XPV{Enter}");
        expect(await answer).toBe("K7QM3XPV");
    });

    it("will not submit half a code", async () => {
        render(<JoinDialog />);
        void ask();
        await userEvent.type(screen.getByTestId("join-code-field"), "K7QM");
        expect(screen.getByTestId("join-code-submit")).toBeDisabled();
    });

    it("will not submit a code with a character ebb never uses", async () => {
        render(<JoinDialog />);
        void ask();
        await userEvent.type(screen.getByTestId("join-code-field"), "K7QM3XPO");
        expect(screen.getByTestId("join-code-submit")).toBeDisabled();
    });

    it("takes the focus, because a debater who opened it is about to type", () => {
        render(<JoinDialog />);
        void ask();
        expect(screen.getByTestId("join-code-field")).toHaveFocus();
    });

    it("answers null when the debater backs out", async () => {
        render(<JoinDialog />);
        const answer = ask();
        act(() => useJoinDialog.getState().close());
        expect(await answer).toBeNull();
    });

    it("answers null when it leaves the tree, so a join is never left waiting", async () => {
        const { unmount } = render(<JoinDialog />);
        const answer = ask();
        unmount();
        expect(await answer).toBeNull();
    });

    it("starts empty on the next opening, with no code left in the field", async () => {
        render(<JoinDialog />);
        const first = ask();
        await userEvent.type(screen.getByTestId("join-code-field"), "K7QM3XPV");
        act(() => useJoinDialog.getState().close());
        await first;
        void ask();
        expect(screen.getByTestId("join-code-field")).toHaveValue("");
    });
});
