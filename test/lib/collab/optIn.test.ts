import { beforeEach, describe, expect, it } from "vitest";

import { seedDoc } from "@/lib/collab/doc";
import { createMemoryNet } from "@/lib/collab/peerLinkMemory";
import { startCollabSession } from "@/lib/collab/session";
import { makeFlowRound } from "@/lib/model/flow";
import { useFlowStore } from "@/lib/store/useFlowStore";

const net = createMemoryNet();
const round = makeFlowRound({});

/**
 * A fully formed session request. Everything a real one carries is present, so
 * the only reason nothing happens below is the switch itself.
 *
 * The scheduler is a no-op canceller: a dial to a peer that is not listening
 * arms a retry on the session's own clock, and a positive control should
 * record the dial without leaving a timer behind.
 */
function deps(dial: string[] = ["sam", "kim"]) {
    return {
        createLink: net.create("alex"),
        roundId: round.id,
        appVersion: "0.11.0",
        doc: () => seedDoc(round),
        apply: () => [],
        dial,
        schedule: () => () => {},
    };
}

beforeEach(() => {
    net.reset();
});

/**
 * The positive control for the suite below. An empty recorder only means "the
 * gate held" if the same request fills it when the gate is open, so these
 * assertions are what give the off tests their meaning.
 */
describe("with shared editing switched on", () => {
    beforeEach(() => {
        useFlowStore.setState({ collabEnabled: true, collabRelayEnabled: true });
    });

    it("binds one endpoint and listens on it", async () => {
        await startCollabSession(deps([]));
        const bound = net.calls.filter((c) => c.op === "create");
        expect(bound).toHaveLength(1);
        expect(bound[0]!.endpointId).toBe("alex");
        expect(net.calls.filter((c) => c.op === "listen").map((c) => c.endpointId)).toEqual([
            "alex",
        ]);
    });

    it("binds with an explicit discovery and relay config", async () => {
        await startCollabSession(deps([]));
        expect(net.calls.find((c) => c.op === "create")!.config).toEqual({
            discovery: "mdns",
            relay: true,
        });
    });

    it("dials every peer the round already knows", async () => {
        await startCollabSession(deps());
        expect(net.calls.filter((c) => c.op === "dial").map((c) => c.endpointId)).toEqual([
            "sam",
            "kim",
        ]);
    });

    it("hands back a session on the endpoint it bound", async () => {
        const session = await startCollabSession(deps([]));
        expect(session?.endpointId).toBe("alex");
    });
});

describe("with shared editing switched off", () => {
    beforeEach(() => {
        useFlowStore.setState({ collabEnabled: false, collabRelayEnabled: true });
    });

    it("binds no endpoint", async () => {
        await startCollabSession(deps());
        expect(net.calls.map((c) => c.op)).not.toContain("listen");
    });

    it("dials no peer, though two are named", async () => {
        await startCollabSession(deps());
        expect(net.calls.map((c) => c.op)).not.toContain("dial");
    });

    it("publishes no discovery record", async () => {
        await startCollabSession(deps());
        expect(net.calls.map((c) => c.op)).not.toContain("create");
    });

    /**
     * The recorder has no relay concept of its own, so what is proven here is
     * narrower than the name: no transport is ever constructed, and a relay is
     * only reachable through one that is.
     */
    it("constructs no transport at all, so no relay is contacted", async () => {
        await startCollabSession(deps());
        expect(net.calls).toEqual([]);
    });

    it("hands back no session at all", async () => {
        expect(await startCollabSession(deps())).toBeNull();
    });
});

/**
 * The other four routes onto the network gate on the same `collabSettings()`
 * call, and each is held to the off case beside its own behavior: `join.ts` in
 * `join.test.ts`, `inviteListener.ts` in `inviteListener.test.ts`,
 * `startForRound` in `runtimeInvites.test.ts`, and `persistReplica` in
 * `persist.test.ts`.
 */
describe("discovery", () => {
    it("never publishes a DNS record, switch on included", async () => {
        useFlowStore.setState({ collabEnabled: true, collabRelayEnabled: true });
        await startCollabSession(deps([]));
        const configs = net.calls.filter((c) => c.op === "create").map((c) => c.config);
        expect(configs).toHaveLength(1);
        expect(configs[0]!.discovery).toBe("mdns");
        expect(Object.values(configs[0]!)).not.toContain("dns");
    });

    it("follows the relay setting", async () => {
        useFlowStore.setState({ collabEnabled: true, collabRelayEnabled: false });
        await startCollabSession(deps([]));
        expect(net.calls.find((c) => c.op === "create")!.config!.relay).toBe(false);
    });
});
