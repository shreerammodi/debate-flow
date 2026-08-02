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

/** What an endpoint actually prints: 64 hex characters. */
const ALEX = "a1e0".repeat(16);
const RIN = "b2f0".repeat(16);

const saved: Contacts = { [ALEX]: { name: "Alex" } };

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
        expect(resolveContacts({ [ALEX]: {} })).toEqual({});
        expect(resolveContacts({ [ALEX]: { name: "   " } })).toEqual({});
    });

    // Every entry a build before this one wrote carries a role, and the field
    // is gone. Reading one costs that entry its role and never the partner it
    // saves, or an upgrade empties the contact table on its first read.
    it("keeps an entry still carrying a legacy role, minus the role", () => {
        for (const role of ["partner", "admin"]) {
            expect(resolveContacts({ [ALEX]: { name: "Alex", role } })).toEqual({
                [ALEX]: { name: "Alex" },
            });
        }
    });

    it("keeps an entry with no role, which is every contact this build writes", () => {
        expect(resolveContacts({ [ALEX]: { name: "Alex" } })).toEqual({
            [ALEX]: { name: "Alex" },
        });
    });

    /**
     * The relay is where this partner was last found, and the only thing that
     * reaches them from another network. It is also a dial target off a
     * hand-editable file, so a scheme somebody chose is dropped - as is a
     * string too long to be an address - which costs the address and leaves
     * the contact reachable in the room.
     */
    it("keeps an https relay and drops anything else, keeping the contact", () => {
        const homed = { [ALEX]: { name: "Alex", relay: "https://r.example/" } };
        expect(resolveContacts(homed)).toEqual(homed);

        const overlong = `https://r.example/${"a".repeat(256)}`;
        for (const relay of ["http://r.example/", "ws://r.example/", "", 7, overlong]) {
            expect(resolveContacts({ [ALEX]: { name: "Alex", relay } })).toEqual({
                [ALEX]: { name: "Alex" },
            });
        }
    });

    it("keeps the good entries beside a bad one", () => {
        const mixed = { [ALEX]: { name: "Alex" }, [RIN]: { name: "" } };
        expect(Object.keys(resolveContacts(mixed))).toEqual([ALEX]);
    });

    it("drops an entry whose key is not a plausible endpoint id", () => {
        expect(resolveContacts({ "": { name: "Alex" } })).toEqual({});
        expect(resolveContacts({ alex: { name: "Alex" } })).toEqual({});
        expect(resolveContacts({ [`${ALEX}!`]: { name: "Alex" } })).toEqual({});
    });

    // A config.toml is hand-editable, and a key in it decides who counts as a
    // saved partner. TOML and JSON both yield "__proto__" as an own key, which
    // a plain object assigns through the setter rather than storing.
    it("assigns nothing through a prototype key", () => {
        const raw: unknown = JSON.parse(
            `{"__proto__":{"name":"Mallory"},` +
                `"constructor":{"name":"Mallory"},` +
                `"${ALEX}":{"name":"Alex"}}`,
        );
        const table = resolveContacts(raw);
        expect(Object.keys(table)).toEqual([ALEX]);
        expect(Object.getPrototypeOf(table)).toBeNull();
    });

    it("builds a table with nothing on its chain to find", () => {
        expect(Object.getPrototypeOf(resolveContacts(saved))).toBeNull();
    });
});

describe("addContact", () => {
    it("adds one, keyed by endpoint id", () => {
        expect(addContact({}, ALEX, { name: "Alex" })).toEqual(saved);
    });

    it("replaces the entry for a peer already saved, rather than duplicating", () => {
        const renamed = addContact(saved, ALEX, { name: "Alexis" });
        expect(Object.keys(renamed)).toHaveLength(1);
        expect(renamed[ALEX]).toEqual({ name: "Alexis" });
    });

    it("does not mutate the table it was given", () => {
        addContact(saved, "other", { name: "Sam" });
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
