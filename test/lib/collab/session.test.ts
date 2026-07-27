import { beforeEach, describe, expect, it } from "vitest";

import { seedDoc } from "@/lib/collab/doc";
import { merge } from "@/lib/collab/merge";
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
