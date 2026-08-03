import { beforeEach, describe, expect, it } from "vitest";

import { seedDoc } from "@/lib/collab/doc";
import { merge } from "@/lib/collab/merge";
import type { PeerLink, PeerLinkConfig } from "@/lib/collab/peerLink";
import { createMemoryNet, memoryRelay } from "@/lib/collab/peerLinkMemory";
import { forgetRoundPeers } from "@/lib/collab/roundPeers";
import { startCollabSession } from "@/lib/collab/session";
import type { CollabDoc } from "@/lib/collab/types";
import { makeFlowRound, type FlowRound } from "@/lib/model/flow";
import { useFlowStore } from "@/lib/store/useFlowStore";

const net = createMemoryNet();
const ALEX = "a".repeat(64);

let shared: FlowRound;

/** A replica the session can read and write, with no grid behind it. */
function side(base: FlowRound) {
    let doc = seedDoc(base);
    return {
        doc: () => doc,
        apply: (incoming: CollabDoc) => {
            const result = merge(doc, incoming);
            doc = result.doc;
            return result.dropped;
        },
    };
}

/**
 * A link that answers about its relay however the test says, so the wait a
 * real endpoint spends on one is a value here rather than a timer. The last
 * answer stands for every later ask.
 *
 * The session asks once as it binds, so the first answer is the one that is
 * already in hand by the time anything shares.
 */
function linkWithRelays(answers: string[]) {
    const base = net.create(ALEX);
    let asked = 0;
    return async (config: PeerLinkConfig): Promise<PeerLink> => {
        const link = await base(config);
        return {
            ...link,
            async relayUrl() {
                const answer = answers[Math.min(asked, answers.length - 1)];
                asked += 1;
                return answer;
            },
        };
    };
}

function open(createLink: (config: PeerLinkConfig) => Promise<PeerLink>) {
    return startCollabSession({
        createLink,
        roundId: shared.id,
        appVersion: "0.11.0",
        ...side(shared),
    });
}

beforeEach(() => {
    net.reset();
    forgetRoundPeers();
    shared = makeFlowRound({});
    useFlowStore.setState({ collabEnabled: true, collabRelayEnabled: true });
});

describe("share, when relaying is on", () => {
    it("puts the relay it homed on into the ticket", async () => {
        const session = await open(linkWithRelays([memoryRelay(ALEX)]));
        const ticket = await session!.share("editor");
        expect(ticket.relayUrl).toBe(memoryRelay(ALEX));
        await session!.stop();
    });

    it("refuses rather than minting a ticket that names no relay", async () => {
        const session = await open(linkWithRelays([""]));
        await expect(session!.share("editor")).rejects.toThrow(/Could not reach a relay/);
        await session!.stop();
    });

    it("leaves no secret armed by a refusal, so a stranger cannot spend one", async () => {
        const session = await open(linkWithRelays(["", "", memoryRelay(ALEX)]));
        await expect(session!.share("editor")).rejects.toThrow();
        // The second attempt mints its own secret. If the refusal had armed
        // one, the first would still be spendable beside it.
        const ticket = await session!.share("editor");
        expect(ticket.secret).toHaveLength(24);
        await session!.stop();
    });

    it("asks the link again after an empty answer, so a retry can succeed", async () => {
        const session = await open(linkWithRelays(["", "", memoryRelay(ALEX)]));
        await expect(session!.share("editor")).rejects.toThrow();
        const ticket = await session!.share("editor");
        expect(ticket.relayUrl).toBe(memoryRelay(ALEX));
        await session!.stop();
    });

    it("keeps the answer it already has rather than asking again", async () => {
        const session = await open(linkWithRelays([memoryRelay(ALEX), ""]));
        expect((await session!.share("editor")).relayUrl).toBe(memoryRelay(ALEX));
        expect((await session!.share("editor")).relayUrl).toBe(memoryRelay(ALEX));
        await session!.stop();
    });
});

describe("share, when relaying is off", () => {
    beforeEach(() => {
        useFlowStore.setState({ collabEnabled: true, collabRelayEnabled: false });
    });

    it("mints a ticket that names no relay, for a round in one room", async () => {
        const session = await open(linkWithRelays([""]));
        const ticket = await session!.share("editor");
        expect(ticket.relayUrl).toBeUndefined();
        expect(ticket.relay).toBe(false);
        await session!.stop();
    });

    it("never asks the link where it is homed, because it cannot be", async () => {
        const asked: string[] = [];
        const base = net.create(ALEX);
        const session = await open(async (config) => {
            const link = await base(config);
            return {
                ...link,
                async relayUrl() {
                    asked.push("relayUrl");
                    return "";
                },
            };
        });
        await session!.share("editor");
        expect(asked).toEqual([]);
        await session!.stop();
    });
});
