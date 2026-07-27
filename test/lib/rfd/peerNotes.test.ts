import { describe, expect, it } from "vitest";

import type { Contacts } from "@/lib/collab/contacts";
import { authoredPeerNotes } from "@/lib/rfd/peerNotes";

const contacts: Contacts = {
    "zz-known": { name: "Sam", role: "partner" },
    "aa-known": { name: "Rae", role: "coach" },
};

describe("authoredPeerNotes", () => {
    it("returns nothing when there is no decision or no peer notes", () => {
        expect(authoredPeerNotes(undefined, contacts)).toEqual([]);
        expect(authoredPeerNotes({ rfd: "mine" }, contacts)).toEqual([]);
        expect(authoredPeerNotes({ peerNotes: {} }, contacts)).toEqual([]);
    });

    it("labels each peer with their contact name", () => {
        const notes = authoredPeerNotes(
            { peerNotes: { "zz-known": "aff on T", "aa-known": "neg on case" } },
            contacts,
        );
        expect(notes).toEqual([
            { endpointId: "aa-known", author: "Rae", text: "neg on case" },
            { endpointId: "zz-known", author: "Sam", text: "aff on T" },
        ]);
    });

    it("falls back to the short EndpointId for an unknown peer", () => {
        const [note] = authoredPeerNotes(
            { peerNotes: { "0123456789abcdef": "dropped the disad" } },
            contacts,
        );
        expect(note.author).toBe("01234567");
    });

    it("orders by EndpointId, not by display name", () => {
        const renamed: Contacts = {
            "aa-known": { name: "Zoe", role: "coach" },
            "zz-known": { name: "Abe", role: "partner" },
        };
        const ids = authoredPeerNotes(
            { peerNotes: { "zz-known": "second", "aa-known": "first" } },
            renamed,
        ).map((n) => n.endpointId);
        expect(ids).toEqual(["aa-known", "zz-known"]);
    });

    it("drops blank and whitespace-only notes", () => {
        const notes = authoredPeerNotes(
            { peerNotes: { "aa-known": "", "zz-known": "   \n  ", "0123456789": "kept" } },
            contacts,
        );
        expect(notes.map((n) => n.endpointId)).toEqual(["0123456789"]);
    });

    it("skips a non-string entry from a hand-edited file", () => {
        const peerNotes = { "aa-known": 42, "zz-known": "kept" } as unknown as Record<
            string,
            string
        >;
        expect(authoredPeerNotes({ peerNotes }, contacts).map((n) => n.text)).toEqual(["kept"]);
    });
});
