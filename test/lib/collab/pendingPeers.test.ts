import { beforeEach, describe, expect, it } from "vitest";

import { seedDoc } from "@/lib/collab/doc";
import { merge } from "@/lib/collab/merge";
import type { PeerConn, PeerLink, PeerLinkConfig } from "@/lib/collab/peerLink";
import { createMemoryNet } from "@/lib/collab/peerLinkMemory";
import { forgetRoundPeers } from "@/lib/collab/roundPeers";
import { startCollabSession, type PendingPeer } from "@/lib/collab/session";
import { encodeTicket } from "@/lib/collab/ticket";
import type { CollabDoc } from "@/lib/collab/types";
import { makeFlowRound, type FlowRound } from "@/lib/model/flow";
import { useFlowStore } from "@/lib/store/useFlowStore";

const net = createMemoryNet();
const ALEX = "a".repeat(64);
const SAM = "b".repeat(64);

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

/** Every delay the session armed, fired on demand rather than by a real timer. */
function clock() {
    const due: (() => void)[] = [];
    return {
        schedule(fn: () => void) {
            due.push(fn);
            return () => {
                const at = due.indexOf(fn);
                if (at >= 0) due.splice(at, 1);
            };
        },
        async tick() {
            for (const fn of due.splice(0, due.length)) fn();
            await settle();
        },
    };
}

async function settle(): Promise<void> {
    for (let i = 0; i < 20; i++) await Promise.resolve();
}

function open(endpointId: string, over: Record<string, unknown> = {}) {
    return startCollabSession({
        createLink: net.create(endpointId),
        roundId: shared.id,
        appVersion: "0.11.0",
        ...side(shared),
        ...over,
    });
}

/**
 * A link that keeps the connections it dialled, so a test can drop one the way
 * a venue's wifi does: closed with no farewell, which is the case the redial
 * ladder exists for.
 */
function tapped(endpointId: string) {
    const base = net.create(endpointId);
    const dialled: PeerConn[] = [];
    return {
        dialled,
        async createLink(config: PeerLinkConfig): Promise<PeerLink> {
            const link = await base(config);
            return {
                ...link,
                async dial(target: string, relayUrl?: string | null) {
                    const conn = await link.dial(target, relayUrl);
                    dialled.push(conn);
                    return conn;
                },
            };
        },
    };
}

beforeEach(() => {
    net.reset();
    forgetRoundPeers();
    shared = makeFlowRound({});
    useFlowStore.setState({ collabEnabled: true, collabRelayEnabled: true });
});

describe("the peers a session is still trying to reach", () => {
    it("is empty for a session nobody has joined", async () => {
        const timer = clock();
        const session = await open(ALEX, { schedule: timer.schedule });
        expect(session!.pending()).toEqual([]);
        await session!.stop();
    });

    it("lists a remembered peer whose opening dial found nobody", async () => {
        const seen: PendingPeer[][] = [];
        const timer = clock();
        const session = await open(ALEX, {
            dial: [SAM],
            schedule: timer.schedule,
            onPendingChanged: (p: PendingPeer[]) => seen.push(p),
        });
        expect(session!.pending()).toEqual([{ endpointId: SAM, unreachable: true }]);
        expect(seen.at(-1)).toEqual([{ endpointId: SAM, unreachable: true }]);
        await session!.stop();
    });

    it("drops a peer from the list the moment it answers", async () => {
        const timer = clock();
        const host = await open(SAM, { schedule: timer.schedule });
        const ticket = encodeTicket(await host!.share("editor"));
        const guest = await open(ALEX, { dial: [SAM], ticket, schedule: timer.schedule });
        await settle();
        expect(guest!.pending()).toEqual([]);
        expect(guest!.peers().map((p) => p.endpointId)).toEqual([SAM]);
        await host!.stop();
        await guest!.stop();
    });

    it("says nothing about a partner who left, because they are not coming back", async () => {
        const timer = clock();
        const host = await open(SAM, { schedule: timer.schedule });
        const ticket = encodeTicket(await host!.share("editor"));
        const guest = await open(ALEX, { dial: [SAM], ticket, schedule: timer.schedule });
        await settle();
        // A window closing says goodbye, which is what stops the far side
        // going back for it.
        await host!.stop();
        await settle();
        expect(guest!.pending()).toEqual([]);
        await guest!.stop();
    });

    it("puts a peer back on the list when its link drops, and does not blame it yet", async () => {
        const timer = clock();
        const host = await open(SAM, { schedule: timer.schedule });
        const ticket = encodeTicket(await host!.share("editor"));
        const tap = tapped(ALEX);
        const guest = await open(ALEX, {
            createLink: tap.createLink,
            dial: [SAM],
            ticket,
            schedule: timer.schedule,
        });
        await settle();
        // A cable pulled, not a window closed: no farewell, so this side goes
        // back for them.
        tap.dialled[0].close();
        await settle();
        expect(guest!.pending()).toEqual([{ endpointId: SAM, unreachable: false }]);
        await host!.stop();
        await guest!.stop();
    });

    it("calls a peer unreachable once a redial has come back", async () => {
        const timer = clock();
        const host = await open(SAM, { schedule: timer.schedule });
        const ticket = encodeTicket(await host!.share("editor"));
        const tap = tapped(ALEX);
        const guest = await open(ALEX, {
            createLink: tap.createLink,
            dial: [SAM],
            ticket,
            schedule: timer.schedule,
        });
        await settle();
        tap.dialled[0].close();
        await settle();
        await host!.stop();
        await settle();
        await timer.tick();
        expect(guest!.pending()).toEqual([{ endpointId: SAM, unreachable: true }]);
        await guest!.stop();
    });

    it("forgets a peer the debater cut loose", async () => {
        const timer = clock();
        const session = await open(ALEX, { dial: [SAM], schedule: timer.schedule });
        session!.disconnect(SAM);
        expect(session!.pending()).toEqual([]);
        await session!.stop();
    });

    it("is empty once the session stops, so nothing is left being waited for", async () => {
        const timer = clock();
        const session = await open(ALEX, { dial: [SAM], schedule: timer.schedule });
        await session!.stop();
        expect(session!.pending()).toEqual([]);
    });
});
