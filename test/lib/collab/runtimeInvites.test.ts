import { beforeEach, describe, expect, it, vi } from "vitest";

import type { PeerLink, PeerLinkConfig } from "@/lib/collab/peerLink";
import type { MemoryNet } from "@/lib/collab/peerLinkMemory";
import type { CollabSession } from "@/lib/collab/session";

interface Corner {
    message: string;
    action?: { label: string; onClick: () => void };
}

const corners: Corner[] = [];

vi.mock("sonner", () => ({
    toast: Object.assign(
        (message: string, opts?: { action?: { label: string; onClick: () => void } }) => {
            corners.push({ message, action: opts?.action });
        },
        {
            warning: () => {},
            error: () => {},
            success: () => {},
            info: () => {},
        },
    ),
}));

/** Filled in below, once the real transport module has been imported. */
const transport = vi.hoisted(() => ({
    link: null as ((config: PeerLinkConfig) => Promise<PeerLink>) | null,
}));

vi.mock("@/lib/collab/peerLink", async (importOriginal) => ({
    ...(await importOriginal<typeof import("@/lib/collab/peerLink")>()),
    createPeerLinkFor: (config: PeerLinkConfig) => transport.link!(config),
}));

import { seedDoc } from "@/lib/collab/doc";
import { createMemoryNet } from "@/lib/collab/peerLinkMemory";
import { clearReplica } from "@/lib/collab/replica";
import { endSession, startForRound, syncInviteWatch } from "@/lib/collab/runtime";
import { startCollabSession } from "@/lib/collab/session";
import { encodeTicket } from "@/lib/collab/ticket";
import { makeFlowRound, type FlowRound } from "@/lib/model/flow";
import { useCollabStore } from "@/lib/store/useCollabStore";
import { useFlowStore } from "@/lib/store/useFlowStore";

const net: MemoryNet = createMemoryNet();
transport.link = net.create("me");

let round: FlowRound;

function listens(): number {
    return net.calls.filter((c) => c.op === "listen").length;
}

beforeEach(async () => {
    await endSession();
    net.reset();
    corners.length = 0;
    clearReplica();
    useCollabStore.getState().reset();
    useFlowStore.setState({
        collabEnabled: true,
        collabRelayEnabled: true,
        shadowMode: false,
        contacts: {},
        docPath: "/flows/round-3-harvard.ebb",
    });
    round = makeFlowRound({});
});

describe("the idle invite listener", () => {
    it("binds nothing while shared editing is off", async () => {
        useFlowStore.setState({ collabEnabled: false });
        await syncInviteWatch();
        expect(net.calls).toEqual([]);
    });

    it("binds one endpoint, however many times it is asked", async () => {
        await syncInviteWatch();
        await syncInviteWatch();
        expect(listens()).toBe(1);
    });

    it("lets go of the endpoint when the switch goes off", async () => {
        await syncInviteWatch();
        useFlowStore.setState({ collabEnabled: false });
        await syncInviteWatch();
        expect(net.calls.filter((c) => c.op === "stop")).toHaveLength(1);
    });

    it("lets go before a session takes the endpoint", async () => {
        await syncInviteWatch();
        await startForRound(round);
        const order = net.calls.filter((c) => c.op === "stop" || c.op === "listen");
        expect(order.map((c) => c.op)).toEqual(["listen", "stop", "listen"]);
    });

    it("takes the endpoint back when the session ends", async () => {
        await startForRound(round);
        await endSession();
        expect(listens()).toBe(2);
    });

    it("publishes the identity a partner is handed, before any round is shared", async () => {
        // reset() keeps a learned identity on purpose, so this clears it.
        useCollabStore.setState({ endpointId: null });
        expect(useCollabStore.getState().endpointId).toBeNull();
        await syncInviteWatch();
        expect(useCollabStore.getState().endpointId).toBeTruthy();
    });

    it("keeps the identity after the endpoint is let go, because it outlives it", async () => {
        await syncInviteWatch();
        const id = useCollabStore.getState().endpointId;
        useFlowStore.setState({ collabEnabled: false });
        await syncInviteWatch();
        expect(useCollabStore.getState().endpointId).toBe(id);
    });
});

describe("a session opened for a round", () => {
    it("comes up for the round that is loaded", async () => {
        const session = await startForRound(round);
        expect(session!.roundId).toBe(round.id);
    });

    it("ends the session it was holding for another round", async () => {
        const first = await startForRound(round);
        const other = makeFlowRound({});
        const second = await startForRound(other);
        expect(second).not.toBe(first);
        expect(second!.roundId).toBe(other.id);
    });

    it("hands back the same session for the round already open", async () => {
        const first = await startForRound(round);
        expect(await startForRound(round)).toBe(first);
    });

    it("leaves no chip behind when the transport will not come up", async () => {
        transport.link = () => Promise.reject(new Error("no shell"));
        await expect(startForRound(round)).rejects.toThrow("no shell");
        expect(useCollabStore.getState().status).toBe("off");
        transport.link = net.create("me");
    });
});

describe("a peer nobody has saved", () => {
    /** Brings a guest onto the host's session, the way a ticket does. */
    async function guestJoins(host: CollabSession): Promise<void> {
        const ticket = encodeTicket(host.share("partner"));
        await startCollabSession({
            createLink: net.create("sam"),
            roundId: round.id,
            appVersion: "0.11.0",
            doc: () => seedDoc(round),
            apply: () => [],
            ticket,
            dial: ["me"],
        });
        for (let i = 0; i < 10; i++) await Promise.resolve();
    }

    it("is offered as a contact, once", async () => {
        const host = await startForRound(round);
        await guestJoins(host!);
        const offers = corners.filter((c) => c.message.startsWith("Save "));
        expect(offers).toHaveLength(1);
        expect(offers[0].message).toBe("Save sam as a partner?");
    });

    it("is saved by the one click, under a name the debater can change", async () => {
        const host = await startForRound(round);
        await guestJoins(host!);
        corners.find((c) => c.action)!.action!.onClick();
        expect(useFlowStore.getState().contacts.sam).toEqual({ name: "sam", role: "partner" });
    });

    it("is not offered again once they are saved", async () => {
        useFlowStore.setState({ contacts: { sam: { name: "Sam", role: "partner" } } });
        const host = await startForRound(round);
        await guestJoins(host!);
        expect(corners.filter((c) => c.message.startsWith("Save "))).toEqual([]);
    });

    it("is named in the chip by the contact table, not by their id", async () => {
        useFlowStore.setState({ contacts: { sam: { name: "Sam", role: "partner" } } });
        const host = await startForRound(round);
        await guestJoins(host!);
        expect(useCollabStore.getState().peers.map((p) => p.name)).toEqual(["Sam"]);
    });
});
