/**
 * ContactList component tests.
 *
 * Uses the real Zustand store: a contact is only ever edited through it, and
 * the whole table is replaced on every keystroke, so the rows the edit did not
 * touch are as much of the assertion as the row it did.
 */

import { render, screen, within } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { beforeEach, describe, expect, it } from "vitest";

import ContactList from "@/components/collab/ContactList";
import type { Contacts } from "@/lib/collab/contacts";
import { useFlowStore } from "@/lib/store/useFlowStore";

const ALEX = "alexendpointid0000000000";
const RIN = "rinendpointid00000000000";

const TWO: Contacts = {
    [ALEX]: { name: "Alex", role: "partner" },
    [RIN]: { name: "Rin", role: "coach" },
};

beforeEach(() => {
    window.localStorage.clear();
    useFlowStore.setState({ contacts: {} });
});

describe("ContactList", () => {
    it("says where a contact comes from when none is saved", () => {
        render(<ContactList />);
        const empty = screen.getByTestId("contact-list-empty");
        expect(empty).toHaveTextContent("No partners saved yet");
        expect(empty).toHaveTextContent("after a shared flow session");
    });

    it("renames one contact and leaves the other alone", async () => {
        useFlowStore.setState({ contacts: TWO });
        render(<ContactList />);

        const field = screen.getByTestId(`contact-name-${ALEX}`);
        await userEvent.clear(field);
        await userEvent.type(field, "Alexis");

        expect(useFlowStore.getState().contacts).toEqual({
            [ALEX]: { name: "Alexis", role: "partner" },
            [RIN]: { name: "Rin", role: "coach" },
        });
    });

    it("re-roles a contact from can edit to view only", async () => {
        useFlowStore.setState({ contacts: TWO });
        render(<ContactList />);

        const role = screen.getByTestId(`contact-role-${ALEX}`);
        expect(role).toHaveTextContent("can edit");

        await userEvent.click(role);
        await userEvent.click(await screen.findByTestId(`contact-role-${ALEX}-coach`));

        expect(useFlowStore.getState().contacts[ALEX]).toEqual({ name: "Alex", role: "coach" });
        expect(useFlowStore.getState().contacts[RIN]).toEqual({ name: "Rin", role: "coach" });
    });

    it("removes exactly the contact asked for", async () => {
        useFlowStore.setState({ contacts: TWO });
        render(<ContactList />);

        await userEvent.click(screen.getByTestId(`contact-remove-${ALEX}`));

        expect(useFlowStore.getState().contacts).toEqual({ [RIN]: { name: "Rin", role: "coach" } });
        expect(screen.queryByTestId(`contact-row-${ALEX}`)).toBeNull();
        expect(screen.getByTestId(`contact-row-${RIN}`)).toBeInTheDocument();
    });

    it("shows the first eight characters of the EndpointId, and the whole one on hover", () => {
        useFlowStore.setState({ contacts: TWO });
        render(<ContactList />);

        const row = screen.getByTestId(`contact-row-${ALEX}`);
        expect(within(row).getByTitle(ALEX)).toHaveTextContent(ALEX.slice(0, 8));
    });

    // A nameless entry is dropped when the config file is read back, which
    // would take the EndpointId with it.
    it("keeps a peer whose name is emptied, falling back to the short id", async () => {
        useFlowStore.setState({ contacts: TWO });
        render(<ContactList />);

        await userEvent.clear(screen.getByTestId(`contact-name-${ALEX}`));
        expect(useFlowStore.getState().contacts[ALEX].name).toBe("");

        await userEvent.tab();
        expect(useFlowStore.getState().contacts[ALEX].name).toBe(ALEX.slice(0, 8));
    });
});
