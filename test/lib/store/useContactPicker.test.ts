import { beforeEach, describe, expect, it } from "vitest";

import type { Contacts } from "@/lib/collab/contacts";
import { chooseContact, useContactPicker } from "@/lib/store/useContactPicker";

const CONTACTS: Contacts = {
    alex: { name: "Alex" },
    rin: { name: "Rin" },
};

beforeEach(() => {
    useContactPicker.setState({ contacts: null, role: "editor", resolve: null });
});

describe("useContactPicker", () => {
    it("is closed until someone asks", () => {
        expect(useContactPicker.getState().contacts).toBeNull();
    });

    it("opens on the table it was handed, and resolves with the pick", async () => {
        const picked = chooseContact(CONTACTS, "editor");
        expect(useContactPicker.getState().contacts).toBe(CONTACTS);

        useContactPicker.getState().pick("rin");
        await expect(picked).resolves.toBe("rin");
    });

    // The grant is chosen before the picker opens, so the dialog can say what
    // the click is about to hand over rather than asking a second question.
    it("carries the grant it was opened on", () => {
        void chooseContact(CONTACTS, "viewer");
        expect(useContactPicker.getState().role).toBe("viewer");
    });

    it("closes as it resolves, so nothing is left open behind the answer", async () => {
        const picked = chooseContact(CONTACTS, "editor");
        useContactPicker.getState().pick("alex");
        await picked;

        expect(useContactPicker.getState().contacts).toBeNull();
        expect(useContactPicker.getState().resolve).toBeNull();
    });

    it("resolves null on a cancel", async () => {
        const picked = chooseContact(CONTACTS, "editor");
        useContactPicker.getState().cancel();
        await expect(picked).resolves.toBeNull();
    });

    it("cancels a request still pending rather than stranding its caller", async () => {
        const first = chooseContact(CONTACTS, "editor");
        const second = chooseContact(CONTACTS, "editor");

        await expect(first).resolves.toBeNull();
        useContactPicker.getState().pick("alex");
        await expect(second).resolves.toBe("alex");
    });

    it("has nothing to settle when cancelled while closed", () => {
        expect(() => useContactPicker.getState().cancel()).not.toThrow();
    });
});
