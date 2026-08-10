/**
 * ContactPickerDialog component tests.
 *
 * The dialog is the only surface `chooseContact` has, so what the promise
 * settles to is the assertion in every case. The grant is decided before the
 * picker opens, so the answer is a peer and the grant is only what the title
 * says the click is about to hand over.
 */

import { render, screen, waitFor, within } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { act } from "react";
import { beforeEach, describe, expect, it } from "vitest";

import ContactPickerDialog from "@/components/collab/ContactPickerDialog";
import type { Contacts } from "@/lib/collab/contacts";
import type { Role } from "@/lib/collab/types";
import { chooseContact, useContactPicker } from "@/lib/store/useContactPicker";

const ALEX = "alex-endpoint";
const RIN = "rin-endpoint";
const JAY = "jay-endpoint";

/** Three, so a move by one row is distinguishable from a wrap. */
const THREE: Contacts = {
    [ALEX]: { name: "Alex" },
    [RIN]: { name: "Rin" },
    [JAY]: { name: "Jay" },
};

const PENDING = Symbol("pending");

/** The class the cursor is drawn with, and the only mark that a row is on. */
const CURSOR_CLASS = "bg-accent";

/** Opens the picker the way the invite commands do, from outside React. */
function open(role: Role = "editor", contacts: Contacts = THREE): Promise<string | null> {
    let picked!: Promise<string | null>;
    act(() => {
        picked = chooseContact(contacts, role);
    });
    return picked;
}

/** The row that invites `endpointId`. */
const row = (endpointId: string) => screen.getByTestId(`contact-pick-${endpointId}`);

/**
 * Opens the picker and waits for the list to hold focus. The dialog settles
 * where focus goes a tick after it paints, so a key sent the moment the rows
 * exist lands on the body and reaches no handler at all.
 */
async function openFocused(role: Role = "editor") {
    const picked = open(role);
    await screen.findByTestId("contact-picker");
    const group = screen.getByRole("group", { name: "Saved partners" });
    await waitFor(() => expect(group).toHaveFocus());
    return { picked, group, rows: within(group).getAllByRole("button") };
}

beforeEach(() => {
    useContactPicker.setState({ contacts: null, role: "editor", resolve: null });
});

describe("ContactPickerDialog", () => {
    it("stays out of the way until a command asks who to invite", () => {
        render(<ContactPickerDialog />);
        expect(screen.queryByTestId("contact-picker")).toBeNull();
    });

    it("offers every saved contact, once each", async () => {
        render(<ContactPickerDialog />);
        const picked = open();

        expect(await screen.findByTestId(`contact-pick-${ALEX}`)).toHaveTextContent("Alex");
        expect(row(RIN)).toBeInTheDocument();
        expect(row(JAY)).toBeInTheDocument();

        await userEvent.keyboard("{Escape}");
        await picked;
    });

    // The grant is not chosen here, so the only place a debater can read what
    // the next click hands over is the title and the rows themselves.
    it("says which grant the click is about to hand over", async () => {
        render(<ContactPickerDialog />);
        const editing = open("editor");
        expect(await screen.findByText("Invite a partner to edit")).toBeInTheDocument();
        expect(row(ALEX)).toHaveAccessibleName("Invite Alex to edit");
        await userEvent.keyboard("{Escape}");
        await editing;

        const viewing = open("viewer");
        expect(await screen.findByText("Invite a partner to view")).toBeInTheDocument();
        expect(row(ALEX)).toHaveAccessibleName("Invite Alex to view");
        await userEvent.keyboard("{Escape}");
        await viewing;
    });

    it("answers with the partner that was clicked", async () => {
        render(<ContactPickerDialog />);
        const picked = open();

        await userEvent.click(await screen.findByTestId(`contact-pick-${RIN}`));

        await expect(picked).resolves.toBe(RIN);
        expect(screen.queryByTestId("contact-picker")).toBeNull();
    });

    // This gesture admits a peer to the round, so nobody is the answer until
    // the debater says who. The list itself takes focus and no row is marked,
    // so an Enter that arrives before a choice invites nobody; opening onto a
    // focused row would admit whoever happens to be first.
    it("opens with nobody chosen, so an early Enter invites nobody", async () => {
        render(<ContactPickerDialog />);
        const { picked, rows } = await openFocused();

        for (const button of rows) {
            expect(button).not.toHaveFocus();
            expect(button.className).not.toContain(CURSOR_CLASS);
        }

        await userEvent.keyboard("{Enter}");

        // A race against an already-settled sentinel: PENDING wins only while
        // the request itself has not settled.
        expect(await Promise.race([picked, Promise.resolve(PENDING)])).toBe(PENDING);
        expect(screen.getByTestId("contact-picker")).toBeInTheDocument();

        await userEvent.keyboard("{Escape}");
        await picked;
    });

    it("enters the list at the first contact when the cursor comes down into it", async () => {
        render(<ContactPickerDialog />);
        const { picked } = await openFocused();

        await userEvent.keyboard("{ArrowDown}");

        expect(row(ALEX)).toHaveFocus();
        await userEvent.keyboard("{Enter}");
        await expect(picked).resolves.toBe(ALEX);
    });

    it("enters the list at the last contact when the cursor comes up into it", async () => {
        render(<ContactPickerDialog />);
        const { picked } = await openFocused();

        await userEvent.keyboard("{ArrowUp}");

        expect(row(JAY)).toHaveFocus();
        await userEvent.keyboard("{Enter}");
        await expect(picked).resolves.toBe(JAY);
    });

    it("steps one contact at a time", async () => {
        render(<ContactPickerDialog />);
        const { picked } = await openFocused();

        await userEvent.keyboard("{ArrowDown}{ArrowDown}");

        expect(row(RIN)).toHaveFocus();
        await userEvent.keyboard("{Enter}");
        await expect(picked).resolves.toBe(RIN);
    });

    it("wraps past the last contact back to the first", async () => {
        render(<ContactPickerDialog />);
        const { picked } = await openFocused();

        await userEvent.keyboard("{ArrowDown}{ArrowDown}{ArrowDown}{ArrowDown}{Enter}");

        await expect(picked).resolves.toBe(ALEX);
    });

    it("wraps backwards from the first contact to the last", async () => {
        render(<ContactPickerDialog />);
        const { picked } = await openFocused();

        await userEvent.keyboard("{ArrowDown}{ArrowUp}{Enter}");

        await expect(picked).resolves.toBe(JAY);
    });

    it("dials nobody when dismissed with Escape", async () => {
        render(<ContactPickerDialog />);
        const picked = open();
        await screen.findByTestId("contact-picker");

        await userEvent.keyboard("{Escape}");

        await expect(picked).resolves.toBeNull();
    });

    it("settles a pending request when it leaves the tree", async () => {
        const { unmount } = render(<ContactPickerDialog />);
        const picked = open();
        await screen.findByTestId("contact-picker");

        unmount();

        await expect(picked).resolves.toBeNull();
    });
});
