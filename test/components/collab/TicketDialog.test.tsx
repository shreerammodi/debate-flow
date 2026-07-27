/**
 * TicketDialog component tests.
 *
 * The dialog is the whole route a ticket travels, so what reaches the screen
 * and what the join promise settles to are the assertions.
 */

import { render, screen } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { act } from "react";
import { beforeEach, describe, expect, it, vi } from "vitest";

import TicketDialog from "@/components/collab/TicketDialog";
import { askForTicket, showTicket, useTicketDialog } from "@/lib/store/useTicketDialog";

const { toastSuccess, toastError } = vi.hoisted(() => ({
    toastSuccess: vi.fn(),
    toastError: vi.fn(),
}));

vi.mock("sonner", () => ({ toast: { success: toastSuccess, error: toastError } }));

const TICKET = "ebb1:eyJlbmRwb2ludElkIjoiYWxleCJ9";

beforeEach(() => {
    toastSuccess.mockClear();
    toastError.mockClear();
    useTicketDialog.setState({ open: false, mode: "show", ticket: "", resolve: null });
});

describe("TicketDialog", () => {
    it("stays out of the way until a command has a ticket to move", () => {
        render(<TicketDialog />);
        expect(screen.queryByTestId("ticket-dialog")).toBeNull();
    });

    it("shows a minted ticket as text, not a field to click into", () => {
        render(<TicketDialog />);
        act(() => showTicket(TICKET));

        const text = screen.getByTestId("ticket-text");
        expect(text.textContent).toBe(TICKET);
        expect(text.querySelector("input, textarea")).toBeNull();
    });

    it("copies from inside the click, then says so and gets out of the way", async () => {
        const writeText = vi.fn(async () => {});
        vi.stubGlobal("navigator", { ...navigator, clipboard: { writeText } });
        render(<TicketDialog />);
        act(() => showTicket(TICKET));

        await userEvent.click(screen.getByTestId("ticket-copy"));
        expect(writeText).toHaveBeenCalledWith(TICKET);
        expect(toastSuccess.mock.calls[0]?.[0]).toMatch(/copied/i);
        expect(screen.queryByTestId("ticket-dialog")).toBeNull();
        vi.unstubAllGlobals();
    });

    it("stays open and selects the ticket when the webview refuses the write", async () => {
        const writeText = vi.fn(async () => {
            throw new Error("The request is not allowed by the user agent");
        });
        vi.stubGlobal("navigator", { ...navigator, clipboard: { writeText } });
        render(<TicketDialog />);
        act(() => showTicket(TICKET));

        await userEvent.click(screen.getByTestId("ticket-copy"));
        expect(toastError.mock.calls[0]?.[0]).toMatch(/Cmd\+C/);
        expect(screen.getByTestId("ticket-dialog")).toBeTruthy();
        expect(window.getSelection()?.toString()).toBe(TICKET);
        vi.unstubAllGlobals();
    });

    it("keeps the ticket on screen while a closed share animates out", async () => {
        const writeText = vi.fn(async () => {});
        vi.stubGlobal("navigator", { ...navigator, clipboard: { writeText } });
        render(<TicketDialog />);
        act(() => showTicket(TICKET));

        await userEvent.click(screen.getByTestId("ticket-copy"));
        // The dialog outlives `open` by the length of its exit animation, so a
        // mode that flipped here would show the join field on the way out.
        expect(useTicketDialog.getState().open).toBe(false);
        expect(useTicketDialog.getState().mode).toBe("show");
        expect(useTicketDialog.getState().ticket).toBe(TICKET);
        vi.unstubAllGlobals();
    });

    it("settles a join with the ticket that was pasted in", async () => {
        render(<TicketDialog />);
        let asked!: Promise<string | null>;
        act(() => {
            asked = askForTicket();
        });

        await userEvent.type(screen.getByTestId("ticket-input"), `  ${TICKET}  `);
        await userEvent.click(screen.getByTestId("ticket-submit"));
        expect(await asked).toBe(TICKET);
    });

    it("refuses to join on an empty field", async () => {
        render(<TicketDialog />);
        act(() => {
            void askForTicket();
        });
        expect(screen.getByTestId("ticket-submit")).toBeDisabled();
    });

    it("settles as a cancel when dismissed", async () => {
        render(<TicketDialog />);
        let asked!: Promise<string | null>;
        act(() => {
            asked = askForTicket();
        });

        await userEvent.keyboard("{Escape}");
        expect(await asked).toBeNull();
    });

    it("settles a pending join when it leaves the tree", async () => {
        const { unmount } = render(<TicketDialog />);
        let asked!: Promise<string | null>;
        act(() => {
            asked = askForTicket();
        });

        unmount();
        expect(await asked).toBeNull();
    });

    it("does not strand a join when a share opens over it", async () => {
        render(<TicketDialog />);
        let asked!: Promise<string | null>;
        act(() => {
            asked = askForTicket();
        });

        act(() => showTicket(TICKET));
        expect(await asked).toBeNull();
        expect(screen.getByTestId("ticket-text")).toBeTruthy();
    });
});
