import { beforeEach, describe, expect, it, vi } from "vitest";

import type { Contacts } from "@/lib/collab/contacts";
import { inviteFrom, inviteToastFor, shouldAnnounceInvite } from "@/lib/collab/invite";
import { PROTOCOL_MAJOR, type WireMessage } from "@/lib/collab/peerLink";

const ALEX = "k51qzi5uqu5dlalex";
const STRANGER = "k51qzi5uqu5dlwho";

const contacts: Contacts = { [ALEX]: { name: "Alex" } };

describe("shouldAnnounceInvite", () => {
    it("announces an invite from a saved contact", () => {
        expect(shouldAnnounceInvite(contacts, ALEX)).toBe(true);
    });

    it("says nothing at all for a peer nobody saved", () => {
        // Otherwise anyone who learns an EndpointId can put a toast on a
        // debater's screen mid-round.
        expect(shouldAnnounceInvite(contacts, STRANGER)).toBe(false);
    });

    it("says nothing when no contacts are saved at all", () => {
        expect(shouldAnnounceInvite({}, ALEX)).toBe(false);
    });
});

describe("inviteToastFor", () => {
    it("names the person and the round", () => {
        expect(inviteToastFor(contacts, ALEX, "Round 3 - Harvard")).toBe(
            "Alex shared Round 3 - Harvard",
        );
    });

    it("falls back to a short id when the round has no name", () => {
        expect(inviteToastFor(contacts, ALEX, "")).toBe("Alex shared a round");
    });
});

function hello(from: string, roundId: string, label?: string): WireMessage {
    return {
        type: "hello",
        protocol: PROTOCOL_MAJOR,
        app: "0.11.0",
        endpointId: from,
        roundId,
        role: "editor",
        capabilities: [],
        ...(label === undefined ? {} : { label }),
    };
}

describe("inviteFrom", () => {
    it("reads a contact's dial about another round as an invitation", () => {
        expect(inviteFrom(hello(ALEX, "r2", "Round 3"), contacts, "r1", ALEX)).toEqual({
            endpointId: ALEX,
            roundId: "r2",
            label: "Round 3",
        });
    });

    it("takes an invitation with no round open at all", () => {
        expect(inviteFrom(hello(ALEX, "r2"), contacts, null, ALEX)?.label).toBe("");
    });

    it("is silent for a peer nobody saved", () => {
        expect(inviteFrom(hello(STRANGER, "r2"), contacts, "r1", STRANGER)).toBeNull();
    });

    // An EndpointId is public: it ships inside every ticket its owner hands
    // out. Naming a saved contact must not borrow their standing.
    it("is silent for a stranger dialling under a saved contact's id", () => {
        expect(inviteFrom(hello(ALEX, "r2", "Round 3"), contacts, "r1", STRANGER)).toBeNull();
    });

    // A key that is on nobody's contact list, including through the prototype.
    it("is silent for a dialler naming a property every object has", () => {
        for (const id of ["constructor", "__proto__", "toString"]) {
            expect(inviteFrom(hello(id, "r2", "join me"), contacts, "r1", id)).toBeNull();
        }
    });

    it("is silent about the round this side is already holding", () => {
        // That dial is a peer joining, which admission answers, not an offer.
        expect(inviteFrom(hello(ALEX, "r1"), contacts, "r1", ALEX)).toBeNull();
    });

    it("is silent across a protocol skew, which has its own refusal", () => {
        const skewed = { ...hello(ALEX, "r2"), protocol: PROTOCOL_MAJOR + 1 };
        expect(inviteFrom(skewed, contacts, "r1", ALEX)).toBeNull();
    });

    it("is silent for anything that is not a hello", () => {
        expect(inviteFrom({ type: "bye" }, contacts, "r1", ALEX)).toBeNull();
    });

    // The label is a contact's string and the corner renders it, so the corner
    // decides how much of one it will take.
    it("takes only as much of a label as a corner can say", () => {
        const notice = inviteFrom(hello(ALEX, "r2", "R".repeat(500)), contacts, "r1", ALEX);
        expect(notice!.label).toBe("R".repeat(120));
    });
});
