import { beforeEach, describe, expect, it } from "vitest";

import type { PeerConn } from "@/lib/collab/peerLink";
import { createMemoryNet } from "@/lib/collab/peerLinkMemory";
import { startCollabSession } from "@/lib/collab/session";
import { useFlowStore } from "@/lib/store/useFlowStore";

const net = createMemoryNet();

beforeEach(() => {
    net.reset();
    useFlowStore.setState({ collabEnabled: true, collabRelayEnabled: true });
});

describe("startCollabSession", () => {
    it("listens on the local endpoint", async () => {
        const session = await startCollabSession({ createLink: net.create("alex") });
        expect(session!.endpointId).toBe("alex");
        expect(net.calls.map((c) => c.op)).toContain("listen");
    });

    it("re-dials the round's peers with no interaction", async () => {
        const host = await startCollabSession({ createLink: net.create("alex") });
        const guest = await startCollabSession({ createLink: net.create("sam"), peers: ["alex"] });
        expect(guest!.peers.map((p) => p.id)).toEqual(["alex"]);
        expect(host!.peers.map((p) => p.id)).toEqual(["sam"]);
    });

    it("keeps running when one peer cannot be reached", async () => {
        const session = await startCollabSession({
            createLink: net.create("alex"),
            peers: ["gone"],
        });
        expect(session).not.toBeNull();
        expect(session!.peers).toEqual([]);
    });

    it("reports an inbound peer to the caller", async () => {
        const seen: PeerConn[] = [];
        await startCollabSession({ createLink: net.create("alex"), onPeer: (p) => seen.push(p) });
        await startCollabSession({ createLink: net.create("sam"), peers: ["alex"] });
        expect(seen.map((p) => p.id)).toEqual(["sam"]);
    });

    it("drops a peer from the list when its link closes", async () => {
        const host = await startCollabSession({ createLink: net.create("alex") });
        const guest = await startCollabSession({ createLink: net.create("sam"), peers: ["alex"] });
        guest!.peers[0].close();
        expect(host!.peers).toEqual([]);
        expect(guest!.peers).toEqual([]);
    });

    it("stops the link it started", async () => {
        const session = await startCollabSession({ createLink: net.create("alex") });
        await session!.stop();
        expect(net.calls.map((c) => c.op)).toContain("stop");
    });
});
