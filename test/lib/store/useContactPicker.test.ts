import { beforeEach, describe, expect, it } from "vitest";

import type { Contacts } from "@/lib/collab/contacts";
import { chooseContact, useContactPicker } from "@/lib/store/useContactPicker";

const CONTACTS: Contacts = {
    alex: { name: "Alex" },
    rin: { name: "Rin" },
};

beforeEach(() => {
    useContactPicker.setState({ contacts: null, resolve: null });
});

describe("useContactPicker", () => {
    it("is closed until someone asks", () => {
        expect(useContactPicker.getState().contacts).toBeNull();
    });

    it("opens on the table it was handed, and resolves with the pick", async () => {
        const picked = chooseContact(CONTACTS);
        expect(useContactPicker.getState().contacts).toBe(CONTACTS);

        useContactPicker.getState().pick({ endpointId: "rin", role: "viewer" });
        await expect(picked).resolves.toEqual({ endpointId: "rin", role: "viewer" });
    });

    // The grade is half the answer, so it travels with the peer rather than
    // being decided again by whoever reads the picked id.
    it("carries the grade the picker was answered with", async () => {
        const picked = chooseContact(CONTACTS);
        useContactPicker.getState().pick({ endpointId: "alex", role: "editor" });
        await expect(picked).resolves.toEqual({ endpointId: "alex", role: "editor" });
    });

    it("closes as it resolves, so nothing is left open behind the answer", async () => {
        const picked = chooseContact(CONTACTS);
        useContactPicker.getState().pick({ endpointId: "alex", role: "editor" });
        await picked;

        expect(useContactPicker.getState().contacts).toBeNull();
        expect(useContactPicker.getState().resolve).toBeNull();
    });

    it("resolves null on a cancel", async () => {
        const picked = chooseContact(CONTACTS);
        useContactPicker.getState().cancel();
        await expect(picked).resolves.toBeNull();
    });

    it("cancels a request still pending rather than stranding its caller", async () => {
        const first = chooseContact(CONTACTS);
        const second = chooseContact(CONTACTS);

        await expect(first).resolves.toBeNull();
        useContactPicker.getState().pick({ endpointId: "alex", role: "editor" });
        await expect(second).resolves.toEqual({ endpointId: "alex", role: "editor" });
    });

    it("has nothing to settle when cancelled while closed", () => {
        expect(() => useContactPicker.getState().cancel()).not.toThrow();
    });
});
