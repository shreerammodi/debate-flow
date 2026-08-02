/**
 * ContactPickerDialog component tests.
 *
 * The dialog is the only surface `chooseContact` has, so what the promise
 * settles to is the assertion in every case. A contact carries no grade, so
 * the grant is part of the answer and not a lookup the caller does after.
 */

import { render, screen, waitFor, within } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { act } from "react";
import { beforeEach, describe, expect, it } from "vitest";

import ContactPickerDialog from "@/components/collab/ContactPickerDialog";
import type { Contacts } from "@/lib/collab/contacts";
import type { Role } from "@/lib/collab/types";
import { chooseContact, type ContactChoice, useContactPicker } from "@/lib/store/useContactPicker";

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

/** The class the cursor is drawn with, and the only mark that a grant is on. */
const CURSOR_CLASS = "bg-accent";

/** Opens the picker the way `collab.invite` does, from outside React. */
function open(contacts: Contacts = THREE): Promise<ContactChoice | null> {
    let picked!: Promise<ContactChoice | null>;
    act(() => {
        picked = chooseContact(contacts);
    });
    return picked;
}

/** The button that grants `role` to `endpointId`, one of the two on its row. */
const grant = (role: Role, endpointId: string) =>
    screen.getByTestId(`contact-pick-${role}-${endpointId}`);

/**
 * Opens the picker and waits for the list to hold focus. The dialog settles
 * where focus goes a tick after it paints, so a key sent the moment the rows
 * exist lands on the body and reaches no handler at all.
 */
async function openFocused() {
    const picked = open();
    await screen.findByTestId("contact-picker");
    const group = screen.getByRole("group", { name: "Saved partners" });
    await waitFor(() => expect(group).toHaveFocus());
    return { picked, group, grants: within(group).getAllByRole("button") };
}

beforeEach(() => {
    useContactPicker.setState({ contacts: null, resolve: null });
});

describe("ContactPickerDialog", () => {
    it("stays out of the way until a command asks who to invite", () => {
        render(<ContactPickerDialog />);
        expect(screen.queryByTestId("contact-picker")).toBeNull();
    });

    it("offers both grants on every saved contact", async () => {
        render(<ContactPickerDialog />);
        const picked = open();

        expect(await screen.findByTestId(`contact-pick-editor-${ALEX}`)).toHaveTextContent("Edit");
        expect(grant("viewer", ALEX)).toHaveTextContent("View");
        expect(grant("editor", RIN)).toBeInTheDocument();
        expect(grant("viewer", JAY)).toBeInTheDocument();

        await userEvent.keyboard("{Escape}");
        await picked;
    });

    it("grants the round to a partner who may write when Edit is clicked", async () => {
        render(<ContactPickerDialog />);
        const picked = open();

        await userEvent.click(await screen.findByTestId(`contact-pick-editor-${RIN}`));

        await expect(picked).resolves.toEqual({ endpointId: RIN, role: "editor" });
        expect(screen.queryByTestId("contact-picker")).toBeNull();
    });

    it("grants read access only when View is clicked", async () => {
        render(<ContactPickerDialog />);
        const picked = open();

        await userEvent.click(await screen.findByTestId(`contact-pick-viewer-${RIN}`));

        await expect(picked).resolves.toEqual({ endpointId: RIN, role: "viewer" });
    });

    // This gesture decides whether a peer may write into the round, so no
    // grant is the answer until the debater says which one. The list itself
    // takes focus and no button is marked, so an Enter that arrives before a
    // choice invites nobody; opening onto a focused Edit would hand out the
    // wide grant to anyone who presses Enter twice.
    it("opens with no grant chosen, so an early Enter invites nobody", async () => {
        render(<ContactPickerDialog />);
        const { picked, grants } = await openFocused();

        for (const button of grants) {
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

        expect(grant("editor", ALEX)).toHaveFocus();
        await userEvent.keyboard("{Enter}");
        await expect(picked).resolves.toEqual({ endpointId: ALEX, role: "editor" });
    });

    it("enters the list at the last contact when the cursor comes up into it", async () => {
        render(<ContactPickerDialog />);
        const { picked } = await openFocused();

        await userEvent.keyboard("{ArrowUp}");

        expect(grant("editor", JAY)).toHaveFocus();
        await userEvent.keyboard("{Enter}");
        await expect(picked).resolves.toEqual({ endpointId: JAY, role: "editor" });
    });

    it("steps a whole row at a time, so an arrow never changes the grant", async () => {
        render(<ContactPickerDialog />);
        const { picked } = await openFocused();

        await userEvent.keyboard("{ArrowDown}{ArrowDown}");

        expect(grant("editor", RIN)).toHaveFocus();
        await userEvent.keyboard("{Enter}");
        await expect(picked).resolves.toEqual({ endpointId: RIN, role: "editor" });
    });

    it("holds the grant column the cursor is in when it steps down", async () => {
        render(<ContactPickerDialog />);
        const { picked } = await openFocused();
        act(() => grant("viewer", ALEX).focus());

        await userEvent.keyboard("{ArrowDown}");

        expect(grant("viewer", RIN)).toHaveFocus();
        await userEvent.keyboard("{Enter}");
        await expect(picked).resolves.toEqual({ endpointId: RIN, role: "viewer" });
    });

    it("wraps past the last contact back to the first", async () => {
        render(<ContactPickerDialog />);
        const { picked } = await openFocused();

        await userEvent.keyboard("{ArrowDown}{ArrowDown}{ArrowDown}{ArrowDown}{Enter}");

        await expect(picked).resolves.toEqual({ endpointId: ALEX, role: "editor" });
    });

    it("wraps backwards from the first contact to the last", async () => {
        render(<ContactPickerDialog />);
        const { picked } = await openFocused();

        await userEvent.keyboard("{ArrowDown}{ArrowUp}{Enter}");

        await expect(picked).resolves.toEqual({ endpointId: JAY, role: "editor" });
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
