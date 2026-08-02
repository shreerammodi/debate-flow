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
    [ALEX]: { name: "Alex" },
    [RIN]: { name: "Rin" },
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
            [ALEX]: { name: "Alexis" },
            [RIN]: { name: "Rin" },
        });
    });

    // What a partner may do is decided at the invitation, for that round only,
    // so a saved contact carries no grade to set here.
    it("offers no grade on a row at all", () => {
        useFlowStore.setState({ contacts: TWO });
        render(<ContactList />);

        const row = screen.getByTestId(`contact-row-${ALEX}`);
        expect(screen.queryByTestId(`contact-role-${ALEX}`)).toBeNull();
        expect(within(row).queryByRole("combobox")).toBeNull();
    });

    it("removes exactly the contact asked for", async () => {
        useFlowStore.setState({ contacts: TWO });
        render(<ContactList />);

        await userEvent.click(screen.getByTestId(`contact-remove-${ALEX}`));

        expect(useFlowStore.getState().contacts).toEqual({ [RIN]: { name: "Rin" } });
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

describe("adding a partner by hand", () => {
    const ID = "b".repeat(64);

    it("saves a partner nobody has shared with yet", async () => {
        render(<ContactList />);

        await userEvent.type(screen.getByTestId("add-contact-name"), "Rin");
        await userEvent.type(screen.getByTestId("add-contact-id"), ID);
        await userEvent.click(screen.getByTestId("add-contact-save"));

        expect(useFlowStore.getState().contacts[ID]).toEqual({ name: "Rin" });
    });

    it("empties the form so a second partner does not inherit the first", async () => {
        render(<ContactList />);

        await userEvent.type(screen.getByTestId("add-contact-name"), "Rin");
        await userEvent.type(screen.getByTestId("add-contact-id"), ID);
        await userEvent.click(screen.getByTestId("add-contact-save"));

        expect(screen.getByTestId("add-contact-name")).toHaveValue("");
        expect(screen.getByTestId("add-contact-id")).toHaveValue("");
    });

    it("refuses a name with no id, and an id with no name", async () => {
        render(<ContactList />);
        expect(screen.getByTestId("add-contact-save")).toBeDisabled();

        await userEvent.type(screen.getByTestId("add-contact-name"), "Rin");
        expect(screen.getByTestId("add-contact-save")).toBeDisabled();

        await userEvent.clear(screen.getByTestId("add-contact-name"));
        await userEvent.type(screen.getByTestId("add-contact-id"), ID);
        expect(screen.getByTestId("add-contact-save")).toBeDisabled();
    });

    it("says so when what was pasted is not an id at all", async () => {
        render(<ContactList />);

        await userEvent.type(screen.getByTestId("add-contact-name"), "Rin");
        await userEvent.type(screen.getByTestId("add-contact-id"), "ebb1:whatever");
        expect(screen.getByTestId("add-contact-error")).toBeTruthy();
        expect(screen.getByTestId("add-contact-save")).toBeDisabled();
    });

    it("refuses to add a second entry over a partner already saved", async () => {
        useFlowStore.setState({ contacts: { [ID]: { name: "Rin" } } });
        render(<ContactList />);

        await userEvent.type(screen.getByTestId("add-contact-name"), "Rin again");
        await userEvent.type(screen.getByTestId("add-contact-id"), ID);
        expect(screen.getByTestId("add-contact-known")).toHaveTextContent("Rin");
        expect(screen.getByTestId("add-contact-save")).toBeDisabled();
    });
});
