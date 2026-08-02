import { beforeEach, describe, expect, it } from "vitest";

import type { Contacts } from "@/lib/collab/contacts";
import { seedDoc } from "@/lib/collab/doc";
import type { InviteNotice } from "@/lib/collab/invite";
import { createMemoryNet } from "@/lib/collab/peerLinkMemory";
import { startCollabSession, type CollabSession } from "@/lib/collab/session";
import { makeFlowRound, type FlowRound } from "@/lib/model/flow";
import { useFlowStore } from "@/lib/store/useFlowStore";

const net = createMemoryNet();

let alexRound: FlowRound;
let samRound: FlowRound;
let heard: InviteNotice[];

const knowsAlex: Contacts = { alex: { name: "Alex" } };

/** Alex, holding the round they are about to offer. */
async function alexSession(): Promise<CollabSession> {
    return (await startCollabSession({
        createLink: net.create("alex"),
        roundId: alexRound.id,
        roundLabel: "Round 3 - Harvard",
        appVersion: "0.11.0",
        doc: () => seedDoc(alexRound),
        apply: () => [],
    }))!;
}

/** Sam, mid-round on something else, with Alex saved. */
async function samSession(contacts: Contacts = knowsAlex): Promise<CollabSession> {
    return (await startCollabSession({
        createLink: net.create("sam"),
        roundId: samRound.id,
        appVersion: "0.11.0",
        doc: () => seedDoc(samRound),
        apply: () => [],
        contacts: () => contacts,
        onInvite: (notice) => heard.push(notice),
    }))!;
}

async function settle(): Promise<void> {
    for (let i = 0; i < 10; i++) await Promise.resolve();
}

beforeEach(() => {
    net.reset();
    heard = [];
    useFlowStore.setState({ collabEnabled: true, collabRelayEnabled: true });
    alexRound = makeFlowRound({});
    samRound = makeFlowRound({});
});

describe("a contact invited into a round they do not hold", () => {
    it("hears the offer, named by the round the host calls it", async () => {
        const sam = await samSession();
        const alex = await alexSession();
        await alex.invite("sam", "editor");
        await settle();
        expect(heard).toEqual([
            { endpointId: "alex", roundId: alexRound.id, label: "Round 3 - Harvard" },
        ]);
        expect(sam.peers()).toEqual([]);
    });

    it("joins nothing until the debater acts", async () => {
        const sam = await samSession();
        const alex = await alexSession();
        await alex.invite("sam", "editor");
        await settle();
        // Their own round is untouched, and no peer was added on either side.
        expect(sam.peers()).toEqual([]);
        expect(alex.peers()).toEqual([]);
    });

    it("counts as delivered on the host's side, not as a failure", async () => {
        await samSession();
        const alex = await alexSession();
        await expect(alex.invite("sam", "editor")).resolves.toBeUndefined();
    });

    it("says nothing to a session that has not saved the dialler", async () => {
        await samSession({});
        const alex = await alexSession();
        // The ordinary silent refusal, which reaches the dialler as an error
        // and the receiver not at all.
        await expect(alex.invite("sam", "editor")).rejects.toThrow("refused");
        await settle();
        expect(heard).toEqual([]);
    });

    it("raises nothing for the round this side is already holding", async () => {
        // Once both sides are on one round the same dial is a peer joining,
        // which admission answers on its own terms.
        samRound = alexRound;
        await samSession();
        const alex = await alexSession();
        await expect(alex.invite("sam", "editor")).rejects.toThrow("refused");
        await settle();
        expect(heard).toEqual([]);
    });
});
