import { beforeEach, describe, expect, it } from "vitest";

import { seedDoc } from "@/lib/collab/doc";
import { merge } from "@/lib/collab/merge";
import type { PeerConn, PeerLinkFactory } from "@/lib/collab/peerLink";
import { createMemoryNet } from "@/lib/collab/peerLinkMemory";
import { startCollabSession, type CollabPeer } from "@/lib/collab/session";
import { encodeTicket } from "@/lib/collab/ticket";
import type { CollabDoc } from "@/lib/collab/types";
import { makeFlowRound, type FlowRound } from "@/lib/model/flow";
import { useFlowStore } from "@/lib/store/useFlowStore";

const net = createMemoryNet();

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

function open(endpointId: string, over: Record<string, unknown> = {}) {
    return startCollabSession({
        createLink: net.create(endpointId),
        roundId: shared.id,
        appVersion: "0.11.0",
        ...side(shared),
        ...over,
    });
}

async function settle(): Promise<void> {
    for (let i = 0; i < 20; i++) await Promise.resolve();
}

beforeEach(() => {
    net.reset();
    useFlowStore.setState({ collabEnabled: true, collabRelayEnabled: true });
    shared = makeFlowRound({});
});

describe("startCollabSession", () => {
    it("listens on the local endpoint", async () => {
        const session = await open("alex");
        expect(session!.endpointId).toBe("alex");
        expect(session!.roundId).toBe(shared.id);
        expect(net.calls.map((c) => c.op)).toContain("listen");
    });

    it("keeps running when a known peer cannot be reached", async () => {
        const session = await open("alex", { dial: ["gone"] });
        expect(session).not.toBeNull();
        expect(session!.peers()).toEqual([]);
    });

    it("re-dials a known peer with no ticket, which is what resume does", async () => {
        // The host already knows sam, the way a sidecar's peer list says it does.
        const host = await open("alex", { dial: ["sam"] });
        const guest = await open("sam", { dial: ["alex"] });
        await settle();
        expect(guest!.peers().map((p) => p.endpointId)).toEqual(["alex"]);
        expect(host!.peers().map((p) => p.endpointId)).toEqual(["sam"]);
    });

    it("reports the peer list as it changes", async () => {
        const seen: CollabPeer[][] = [];
        const host = await open("alex", {
            dial: ["sam"],
            onPeersChanged: (peers: CollabPeer[]) => seen.push(peers),
        });
        await open("sam", { dial: ["alex"] });
        await settle();
        expect(seen.at(-1)!.map((p) => p.endpointId)).toEqual(["sam"]);
        expect(host!.peers()).toHaveLength(1);
    });

    it("drops a peer from both lists when the link closes", async () => {
        const host = await open("alex", { dial: ["sam"] });
        const guest = await open("sam", { dial: ["alex"] });
        await settle();
        await guest!.stop();
        await settle();
        expect(host!.peers()).toEqual([]);
        expect(guest!.peers()).toEqual([]);
    });

    it("stops the link it started", async () => {
        const session = await open("alex");
        await session!.stop();
        expect(net.calls.map((c) => c.op)).toContain("stop");
    });

    it("mints a ticket that names this host and this round", async () => {
        const session = await open("alex");
        const ticket = session!.share("partner");
        expect(ticket).toMatchObject({
            endpointId: "alex",
            roundId: shared.id,
            role: "partner",
            relay: true,
        });
        expect(encodeTicket(ticket)).toContain("ebb1:");
    });

    it("mints a fresh ticket each time, replacing the unspent one", async () => {
        const session = await open("alex");
        expect(session!.share("partner").secret).not.toBe(session!.share("partner").secret);
    });

    it("carries the relay stance the settings hold into the ticket", async () => {
        useFlowStore.setState({ collabRelayEnabled: false });
        const session = await open("alex");
        expect(session!.share("partner").relay).toBe(false);
    });
});

describe("a link that drops mid-round", () => {
    /** Time the test owns, so a backoff is a step rather than a wait. */
    function manualClock() {
        let pending: { fn: () => void; at: number }[] = [];
        let now = 0;
        return {
            schedule(fn: () => void, ms: number) {
                const entry = { fn, at: now + ms };
                pending.push(entry);
                return () => {
                    pending = pending.filter((p) => p !== entry);
                };
            },
            advance(ms: number) {
                now += ms;
                const due = pending.filter((p) => p.at <= now);
                pending = pending.filter((p) => p.at > now);
                for (const p of due) p.fn();
            },
        };
    }

    /** A link that hands back every connection it dials, so a test can cut one. */
    function watched(endpointId: string, dialled: PeerConn[]): PeerLinkFactory {
        return async (config) => {
            const link = await net.create(endpointId)(config);
            return {
                ...link,
                async dial(target: string, ticket?: string) {
                    const conn = await link.dial(target, ticket);
                    dialled.push(conn);
                    return conn;
                },
            };
        };
    }

    // The dial that opened a link only retried while the session was coming
    // up, and only when a test handed it a scheduler. A wifi blip mid-round
    // left the peer gone for the rest of the round, and the debater's only way
    // back was to close the flow and open it again.
    it("dials the peer again, without anyone being asked", async () => {
        const clock = manualClock();
        const conns: PeerConn[] = [];
        const host = (await open("alex"))!;
        const guest = (await open("sam", {
            createLink: watched("sam", conns),
            ticket: encodeTicket(host.share("partner")),
            dial: ["alex"],
            schedule: clock.schedule,
        }))!;
        await settle();
        expect(guest.peers()).toHaveLength(1);

        conns[0].close();
        await settle();
        expect(guest.peers()).toHaveLength(0);

        // The backoff comes round and the guest reaches the host again, with
        // no ticket and nothing on screen to answer.
        for (let i = 0; i < 6 && guest.peers().length === 0; i++) {
            clock.advance(60_000);
            await settle();
        }
        expect(guest.peers()).toHaveLength(1);
        expect(conns.length).toBeGreaterThan(1);

        await host.stop();
        await guest.stop();
    });

    it("stops trying once the session is over", async () => {
        const clock = manualClock();
        const conns: PeerConn[] = [];
        const host = (await open("alex"))!;
        const guest = (await open("sam", {
            createLink: watched("sam", conns),
            ticket: encodeTicket(host.share("partner")),
            dial: ["alex"],
            schedule: clock.schedule,
        }))!;
        await settle();

        await guest.stop();
        const dialsAtStop = conns.length;
        for (let i = 0; i < 4; i++) {
            clock.advance(60_000);
            await settle();
        }
        expect(conns).toHaveLength(dialsAtStop);
        expect(guest.peers()).toHaveLength(0);

        await host.stop();
    });
});
