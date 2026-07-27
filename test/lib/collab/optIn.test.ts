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
 */
function deps(dial: string[] = ["sam", "kim"]) {
    return {
        createLink: net.create("alex"),
        roundId: round.id,
        appVersion: "0.11.0",
        doc: () => seedDoc(round),
        apply: () => [],
        dial,
    };
}

beforeEach(() => {
    net.reset();
});

describe("with shared editing switched off", () => {
    beforeEach(() => {
        useFlowStore.setState({ collabEnabled: false, collabRelayEnabled: true });
    });

    it("binds no endpoint", async () => {
        await startCollabSession(deps());
        expect(net.calls.filter((c) => c.op === "listen")).toEqual([]);
    });

    it("dials no peer", async () => {
        await startCollabSession(deps());
        expect(net.calls.filter((c) => c.op === "dial")).toEqual([]);
    });

    it("publishes no discovery record", async () => {
        await startCollabSession(deps());
        expect(net.calls.filter((c) => c.op === "create")).toEqual([]);
    });

    it("contacts no relay", async () => {
        await startCollabSession(deps());
        expect(net.calls).toEqual([]);
    });

    it("hands back no session at all", async () => {
        expect(await startCollabSession(deps())).toBeNull();
    });
});

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
