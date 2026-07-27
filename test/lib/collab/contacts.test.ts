import { describe, expect, it } from "vitest";

import {
    addContact,
    contactName,
    isEndpointId,
    isKnown,
    removeContact,
    resolveContacts,
    type Contacts,
} from "@/lib/collab/contacts";

const ALEX = "k51qzi5uqu5dlalexalexalexalexalexalexalexalexalexal";

const saved: Contacts = { [ALEX]: { name: "Alex", role: "partner" } };

describe("resolveContacts", () => {
    it("keeps a well-formed table", () => {
        expect(resolveContacts(saved)).toEqual(saved);
    });

    it("degrades anything that is not a table to no contacts at all", () => {
        for (const raw of [null, undefined, "x", 7, []]) {
            expect(resolveContacts(raw)).toEqual({});
        }
    });

    it("drops an entry with no usable name rather than inventing one", () => {
        expect(resolveContacts({ [ALEX]: { role: "partner" } })).toEqual({});
        expect(resolveContacts({ [ALEX]: { name: "   ", role: "partner" } })).toEqual({});
    });

    it("drops an unknown role instead of defaulting it", () => {
        // Defaulting here would hand edit rights to whatever the file said.
        expect(resolveContacts({ [ALEX]: { name: "Alex", role: "admin" } })).toEqual({});
        expect(resolveContacts({ [ALEX]: { name: "Alex" } })).toEqual({});
    });

    it("keeps a coach as a coach", () => {
        const coach = { [ALEX]: { name: "Coach", role: "coach" } };
        expect(resolveContacts(coach)).toEqual(coach);
    });

    it("keeps the good entries beside a bad one", () => {
        const mixed = { [ALEX]: { name: "Alex", role: "partner" }, bad: { name: "" } };
        expect(Object.keys(resolveContacts(mixed))).toEqual([ALEX]);
    });

    it("drops an entry whose key is not a plausible endpoint id", () => {
        expect(resolveContacts({ "": { name: "Alex", role: "partner" } })).toEqual({});
    });
});

describe("addContact", () => {
    it("adds one, keyed by endpoint id", () => {
        expect(addContact({}, ALEX, { name: "Alex", role: "partner" })).toEqual(saved);
    });

    it("replaces the entry for a peer already saved, rather than duplicating", () => {
        const renamed = addContact(saved, ALEX, { name: "Alexis", role: "coach" });
        expect(Object.keys(renamed)).toHaveLength(1);
        expect(renamed[ALEX]).toEqual({ name: "Alexis", role: "coach" });
    });

    it("does not mutate the table it was given", () => {
        addContact(saved, "other", { name: "Sam", role: "partner" });
        expect(Object.keys(saved)).toEqual([ALEX]);
    });
});

describe("removeContact", () => {
    it("removes by endpoint id", () => {
        expect(removeContact(saved, ALEX)).toEqual({});
    });

    it("is a no-op for a peer that was never saved", () => {
        expect(removeContact(saved, "nobody")).toEqual(saved);
    });
});

describe("contactName", () => {
    it("uses the saved name", () => {
        expect(contactName(saved, ALEX)).toBe("Alex");
    });

    it("falls back to a short form, because an endpoint id is not readable", () => {
        const short = contactName({}, ALEX);
        expect(short.length).toBeLessThan(ALEX.length);
        expect(ALEX.startsWith(short)).toBe(true);
    });

    it("takes the name a peer broadcast over the short form", () => {
        expect(contactName({}, ALEX, "Rin")).toBe("Rin");
    });

    it("keeps the saved name, so a peer cannot rename themselves on your screen", () => {
        expect(contactName(saved, ALEX, "Someone Else")).toBe("Alex");
    });

    it("ignores a broadcast name that is only whitespace", () => {
        expect(contactName({}, ALEX, "   ")).toBe(ALEX.slice(0, 8));
    });
});

describe("isKnown", () => {
    it("separates a saved peer from a stranger", () => {
        expect(isKnown(saved, ALEX)).toBe(true);
        expect(isKnown(saved, "stranger")).toBe(false);
    });
});

describe("isEndpointId", () => {
    it("takes the hex form an endpoint prints", () => {
        expect(isEndpointId("a".repeat(64))).toBe(true);
        expect(isEndpointId("A1B2".repeat(16))).toBe(true);
    });

    it("takes the base32 form iroh also parses", () => {
        expect(isEndpointId("abcdefghijklmnopqrstuvwxyz234567".padEnd(52, "a"))).toBe(true);
    });

    it("rejects a length no key has", () => {
        expect(isEndpointId("")).toBe(false);
        expect(isEndpointId("a".repeat(63))).toBe(false);
        expect(isEndpointId("a".repeat(51))).toBe(false);
    });

    it("rejects a pasted ticket, which is the likely mistake", () => {
        expect(isEndpointId(`ebb1:${"a".repeat(64)}`)).toBe(false);
    });

    it("rejects characters no encoding uses", () => {
        expect(isEndpointId(`${"a".repeat(63)} `)).toBe(false);
        expect(isEndpointId("!".repeat(64))).toBe(false);
    });
});
