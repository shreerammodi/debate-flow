/**
 * ContactPickerDialog component tests.
 *
 * The dialog is the only surface `chooseContact` has, so what the promise
 * settles to is the assertion in every case.
 */

import { render, screen } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { act } from "react";
import { beforeEach, describe, expect, it } from "vitest";

import ContactPickerDialog from "@/components/collab/ContactPickerDialog";
import type { Contacts } from "@/lib/collab/contacts";
import { chooseContact, useContactPicker } from "@/lib/store/useContactPicker";

const ALEX = "alex-endpoint";
const RIN = "rin-endpoint";

const TWO: Contacts = {
    [ALEX]: { name: "Alex", role: "partner" },
    [RIN]: { name: "Rin", role: "coach" },
};

/** Opens the picker the way `collab.invite` does, from outside React. */
function open(contacts: Contacts = TWO): Promise<string | null> {
    let picked!: Promise<string | null>;
    act(() => {
        picked = chooseContact(contacts);
    });
    return picked;
}

beforeEach(() => {
    useContactPicker.setState({ contacts: null, resolve: null });
});

describe("ContactPickerDialog", () => {
    it("stays out of the way until a command asks who to invite", () => {
        render(<ContactPickerDialog />);
        expect(screen.queryByTestId("contact-picker")).toBeNull();
    });

    it("lists every saved contact with what its role may do", async () => {
        render(<ContactPickerDialog />);
        const picked = open();

        expect(await screen.findByTestId(`contact-pick-${ALEX}`)).toHaveTextContent("can edit");
        expect(screen.getByTestId(`contact-pick-${RIN}`)).toHaveTextContent("view only");

        await userEvent.keyboard("{Escape}");
        await picked;
    });

    it("resolves with the contact that was clicked", async () => {
        render(<ContactPickerDialog />);
        const picked = open();

        await userEvent.click(await screen.findByTestId(`contact-pick-${RIN}`));

        await expect(picked).resolves.toBe(RIN);
        expect(screen.queryByTestId("contact-picker")).toBeNull();
    });

    it("walks the list with the arrow keys and picks with Enter", async () => {
        render(<ContactPickerDialog />);
        const picked = open();
        await screen.findByTestId(`contact-pick-${ALEX}`);

        await userEvent.keyboard("{ArrowDown}{Enter}");

        await expect(picked).resolves.toBe(RIN);
    });

    it("wraps from the last contact back to the first", async () => {
        render(<ContactPickerDialog />);
        const picked = open();
        await screen.findByTestId(`contact-pick-${ALEX}`);

        await userEvent.keyboard("{ArrowUp}{ArrowUp}{Enter}");

        await expect(picked).resolves.toBe(ALEX);
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
