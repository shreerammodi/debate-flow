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

const TICKET = "ebb1:eyJlbmRwb2ludElkIjoiYWxleCJ9";

beforeEach(() => {
    useTicketDialog.setState({ showing: null, resolve: null });
});

describe("TicketDialog", () => {
    it("stays out of the way until a command has a ticket to move", () => {
        render(<TicketDialog />);
        expect(screen.queryByTestId("ticket-dialog")).toBeNull();
    });

    it("shows a minted ticket where it can be read and selected", () => {
        render(<TicketDialog />);
        act(() => showTicket(TICKET));

        const field = screen.getByTestId("ticket-text") as HTMLTextAreaElement;
        expect(field.value).toBe(TICKET);
        expect(field.readOnly).toBe(true);
    });

    it("copies from inside the click, which is the only context the webview grants", async () => {
        const writeText = vi.fn(async () => {});
        vi.stubGlobal("navigator", { ...navigator, clipboard: { writeText } });
        render(<TicketDialog />);
        act(() => showTicket(TICKET));

        await userEvent.click(screen.getByTestId("ticket-copy"));
        expect(writeText).toHaveBeenCalledWith(TICKET);
        expect(screen.getByTestId("ticket-copy-hint").textContent).toMatch(/Copied/);
        vi.unstubAllGlobals();
    });

    it("falls back to a manual copy when the webview refuses anyway", async () => {
        const writeText = vi.fn(async () => {
            throw new Error("The request is not allowed by the user agent");
        });
        vi.stubGlobal("navigator", { ...navigator, clipboard: { writeText } });
        render(<TicketDialog />);
        act(() => showTicket(TICKET));

        await userEvent.click(screen.getByTestId("ticket-copy"));
        expect(screen.getByTestId("ticket-copy-hint").textContent).toMatch(/Cmd\+C/);
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
