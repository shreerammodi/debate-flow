import { beforeEach, describe, expect, it, vi } from "vitest";

import type { Contacts } from "@/lib/collab/contacts";
import { inviteToastFor, shouldAnnounceInvite } from "@/lib/collab/invite";

const ALEX = "k51qzi5uqu5dlalex";
const STRANGER = "k51qzi5uqu5dlwho";

const contacts: Contacts = { [ALEX]: { name: "Alex", role: "partner" } };

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
