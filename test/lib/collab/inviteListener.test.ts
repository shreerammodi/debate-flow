import { beforeEach, describe, expect, it } from "vitest";

import type { Contacts } from "@/lib/collab/contacts";
import { helloFrom } from "@/lib/collab/handshake";
import { INVITED, type InviteNotice } from "@/lib/collab/invite";
import { startInviteListener } from "@/lib/collab/inviteListener";
import type { WireMessage } from "@/lib/collab/peerLink";
import { createMemoryNet } from "@/lib/collab/peerLinkMemory";
import { useFlowStore } from "@/lib/store/useFlowStore";

const ALEX = "alex";
const STRANGER = "who";

const net = createMemoryNet();
const contacts: Contacts = { [ALEX]: { name: "Alex", role: "partner" } };

let heard: InviteNotice[];

function listener(table: Contacts = contacts) {
    return startInviteListener({
        createLink: net.create("me"),
        contacts: () => table,
        onInvite: (notice) => heard.push(notice),
    });
}

/** Dials the listener the way a partner offering a round does. */
async function offer(from: string, label: string): Promise<WireMessage[]> {
    const link = await net.create(from)({ discovery: "mdns", relay: true });
    const conn = await link.dial("me");
    const answers: WireMessage[] = [];
    conn.onMessage((msg) => answers.push(msg));
    conn.send(
        helloFrom({
            endpointId: from,
            roundId: "their-round",
            role: "partner",
            appVersion: "0.11.0",
            label,
        }),
    );
    await Promise.resolve();
    return answers;
}

beforeEach(() => {
    net.reset();
    heard = [];
    useFlowStore.setState({
        collabEnabled: true,
        collabRelayEnabled: true,
        collabListenEnabled: true,
    });
});

describe("with shared editing switched off", () => {
    beforeEach(() => {
        useFlowStore.setState({ collabEnabled: false });
    });

    it("binds no endpoint and hands back no listener", async () => {
        expect(await listener()).toBeNull();
        expect(net.calls).toEqual([]);
    });
});

/**
 * Staying bound with no round in hand is the only thing in ebb that reaches
 * the network without a debater asking for a round, so shared editing being
 * available is not enough on its own.
 */
describe("with shared editing on and Listen for invites off", () => {
    beforeEach(() => {
        useFlowStore.setState({ collabListenEnabled: false });
    });

    it("binds no endpoint and hands back no listener", async () => {
        expect(await listener()).toBeNull();
        expect(net.calls).toEqual([]);
    });
});

describe("an idle install", () => {
    it("hears a saved contact offer a round", async () => {
        await listener();
        await offer(ALEX, "Round 3 - Harvard");
        expect(heard).toEqual([
            { endpointId: ALEX, roundId: "their-round", label: "Round 3 - Harvard" },
        ]);
    });

    it("tells the dialler the notice landed, so they stop dialling", async () => {
        await listener();
        const answers = await offer(ALEX, "Round 3");
        expect(answers).toEqual([{ type: "helloAck", ok: false, reason: INVITED }]);
    });

    it("says nothing at all to a peer nobody saved", async () => {
        await listener();
        const answers = await offer(STRANGER, "Round 3");
        expect(heard).toEqual([]);
        expect(answers).toEqual([]);
    });

    it("joins nothing on its own", async () => {
        // The round only lands when the debater says so, so the listener
        // never asks for state.
        await listener();
        const answers = await offer(ALEX, "Round 3");
        expect(answers.some((m) => m.type === "state" || m.type === "vector")).toBe(false);
    });

    it("reaches the network the way a session does, mDNS and no DNS", async () => {
        await listener();
        const config = net.calls.find((c) => c.op === "create")!.config!;
        expect(config.discovery).toBe("mdns");
        expect(Object.values(config)).not.toContain("dns");
    });

    it("follows the relay setting", async () => {
        useFlowStore.setState({ collabRelayEnabled: false });
        await listener();
        expect(net.calls.find((c) => c.op === "create")!.config!.relay).toBe(false);
    });

    it("releases the endpoint when it stops", async () => {
        const held = await listener();
        await held!.stop();
        expect(net.calls.some((c) => c.op === "stop")).toBe(true);
    });

    it("hears nothing more once it has stopped", async () => {
        const held = await listener();
        await held!.stop();
        await offer(ALEX, "Round 3").catch(() => []);
        expect(heard).toEqual([]);
    });
});

describe("the contact table it consults", () => {
    it("is read at the moment of the dial, not at bind", async () => {
        const table: Contacts = {};
        await listener(table);
        table[ALEX] = { name: "Alex", role: "partner" };
        await offer(ALEX, "Round 3");
        expect(heard.map((n) => n.endpointId)).toEqual([ALEX]);
    });
});
