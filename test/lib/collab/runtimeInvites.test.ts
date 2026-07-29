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
import type { InviteNotice } from "@/lib/collab/invite";
import { createMemoryNet } from "@/lib/collab/peerLinkMemory";
import { clearReplica, getReplica } from "@/lib/collab/replica";
import { peerNotePath } from "@/lib/collab/rfdSync";
import { rememberRoundPeers } from "@/lib/collab/roundPeers";
import {
    currentSession,
    endSession,
    inviteContact,
    notifyLocalChange,
    resumeSession,
    startForRound,
    syncInviteWatch,
} from "@/lib/collab/runtime";
import { startCollabSession } from "@/lib/collab/session";
import { encodeTicket } from "@/lib/collab/ticket";
import { makeFlowRound, type FlowRound } from "@/lib/model/flow";
import { useCollabStore } from "@/lib/store/useCollabStore";
import { useFlowStore } from "@/lib/store/useFlowStore";

const net: MemoryNet = createMemoryNet();
/** What iroh hands back. A ticket names the host, so the host holds a real one. */
const ME = "e".repeat(64);
transport.link = net.create(ME);

let round: FlowRound;

function listens(): number {
    return net.calls.filter((c) => c.op === "listen").length;
}

beforeEach(async () => {
    // startForRound asks collabLive(), so every route in this file needs a
    // shell to be offered at all. isDesktop() reads this global.
    (window as unknown as Record<string, unknown>).__TAURI_INTERNALS__ = {};
    await endSession();
    net.reset();
    corners.length = 0;
    clearReplica();
    useCollabStore.getState().reset();
    useFlowStore.setState({
        collabEnabled: true,
        collabRelayEnabled: true,
        collabListenEnabled: true,
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

    // Shared editing being available is not a reason to be on the network:
    // the master switch unlocks Share and Join, and this switch is what puts
    // an endpoint up with no round in hand.
    it("binds nothing while Listen for invites is off", async () => {
        useFlowStore.setState({ collabListenEnabled: false });
        await syncInviteWatch();
        expect(net.calls).toEqual([]);
    });

    it("lets go of the endpoint when Listen for invites goes off", async () => {
        await syncInviteWatch();
        useFlowStore.setState({ collabListenEnabled: false });
        await syncInviteWatch();
        expect(net.calls.filter((c) => c.op === "stop")).toHaveLength(1);
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

    // There is one endpoint per install. A listener that bound while a session
    // was coming up would share it, hear the session's own peers arrive as
    // diallers, and hang up on them - after which every send is refused and
    // the chip still says Connected.
    it("binds nothing while a session is coming up", async () => {
        const start = startForRound(round);
        await syncInviteWatch();
        await start;
        expect(listens()).toBe(1);
    });

    it("lets go of a listener whose stop the shell refuses", async () => {
        // beforeEach ends a session, which binds a listener again, so this
        // starts from nothing bound before counting.
        useFlowStore.setState({ collabEnabled: false });
        await syncInviteWatch();
        net.reset();

        transport.link = async (config) => {
            const link = await net.create(ME)(config);
            return { ...link, stop: () => Promise.reject(new Error("no shell")) };
        };
        useFlowStore.setState({ collabEnabled: true });
        await syncInviteWatch();
        useFlowStore.setState({ collabEnabled: false });
        await expect(syncInviteWatch()).resolves.toBeUndefined();

        // The handle is gone, so turning the switch back on binds a new one
        // rather than believing it still holds the old.
        transport.link = net.create(ME);
        useFlowStore.setState({ collabEnabled: true });
        await syncInviteWatch();
        expect(listens()).toBe(2);
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
        transport.link = net.create(ME);
    });

    it("ends a session whose link will not stop, because End session cannot fail", async () => {
        transport.link = async (config) => {
            const link = await net.create(ME)(config);
            return { ...link, stop: () => Promise.reject(new Error("no shell")) };
        };
        await startForRound(round);
        await expect(endSession()).resolves.toBeUndefined();
        expect(useCollabStore.getState().status).toBe("off");
        transport.link = net.create(ME);
    });

    // An earlier session left this machine's own note in the file under this
    // machine's own id, where the RFD drawer read it back as a partner's.
    it("sheds this machine's own note from the round it opens", async () => {
        const mine = peerNotePath(ME);
        round.scouting.decision = { rfd: "123123", peerNotes: { [ME]: "blah blah blah" } };
        useFlowStore.getState().loadRound(round);
        expect(getReplica()!.round[mine]).toBeDefined();

        await startForRound(round);

        expect(getReplica()!.round[mine]).toBeUndefined();
        const decision = useFlowStore.getState().round!.scouting.decision;
        expect(decision?.peerNotes?.[ME]).toBeUndefined();
        expect(decision?.rfd).toBe("123123");
    });

    it("keeps a real partner's note when it sheds its own", async () => {
        round.scouting.decision = {
            rfd: "mine",
            peerNotes: { [ME]: "echo of mine", sam: "voting neg" },
        };
        useFlowStore.getState().loadRound(round);

        await startForRound(round);

        expect(useFlowStore.getState().round!.scouting.decision?.peerNotes).toEqual({
            sam: "voting neg",
        });
    });
});

describe("a peer nobody has saved", () => {
    /** Brings a guest onto the host's session, the way a ticket does. */
    async function guestJoins(host: CollabSession, displayName?: string): Promise<void> {
        const ticket = encodeTicket(host.share("partner"));
        await startCollabSession({
            createLink: net.create("sam"),
            roundId: round.id,
            appVersion: "0.11.0",
            doc: () => seedDoc(round),
            apply: () => [],
            displayName,
            ticket,
            dial: [ME],
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

    // A join greets the host before the round's own session re-dials, and an
    // older build greets with no name at all. The offer carries the name it
    // will save, so the first one must not be the last word.
    it("is offered again under the name a nameless greeting later supplies", async () => {
        const host = await startForRound(round);
        await guestJoins(host!);
        expect(corners.at(-1)!.message).toBe("Save sam as a partner?");

        await guestJoins(host!, "Rin");
        const offers = corners.filter((c) => c.message.startsWith("Save "));
        expect(offers.at(-1)!.message).toBe("Save Rin as a partner?");
        offers.at(-1)!.action!.onClick();
        expect(useFlowStore.getState().contacts.sam).toEqual({ name: "Rin", role: "partner" });
    });

    it("is not asked about twice while their name stays the same", async () => {
        const host = await startForRound(round);
        await guestJoins(host!, "Rin");
        await guestJoins(host!, "Rin");
        expect(corners.filter((c) => c.message.startsWith("Save "))).toHaveLength(1);
    });

    it("is named in the chip by the contact table, not by their id", async () => {
        useFlowStore.setState({ contacts: { sam: { name: "Sam", role: "partner" } } });
        const host = await startForRound(round);
        await guestJoins(host!);
        expect(useCollabStore.getState().peers.map((p) => p.name)).toEqual(["Sam"]);
    });

    it("is offered under the name they broadcast, not their id", async () => {
        const host = await startForRound(round);
        await guestJoins(host!, "Rin");
        expect(corners.find((c) => c.message.startsWith("Save "))!.message).toBe(
            "Save Rin as a partner?",
        );
        corners.find((c) => c.action)!.action!.onClick();
        expect(useFlowStore.getState().contacts.sam).toEqual({ name: "Rin", role: "partner" });
    });

    it("is named in the chip by what they broadcast until they are saved", async () => {
        const host = await startForRound(round);
        await guestJoins(host!, "Rin");
        expect(useCollabStore.getState().peers.map((p) => p.name)).toEqual(["Rin"]);
    });

    it("keeps a saved name over the one they broadcast", async () => {
        useFlowStore.setState({ contacts: { sam: { name: "Sam", role: "partner" } } });
        const host = await startForRound(round);
        await guestJoins(host!, "Rin");
        expect(useCollabStore.getState().peers.map((p) => p.name)).toEqual(["Sam"]);
    });
});

describe("sharing a round with a saved contact", () => {
    /** Sam, mid-round on something else, with this machine saved. */
    async function samMidRound(heard: InviteNotice[]): Promise<void> {
        await startCollabSession({
            createLink: net.create("sam"),
            roundId: makeFlowRound({}).id,
            appVersion: "0.11.0",
            doc: () => seedDoc(round),
            apply: () => [],
            contacts: () => ({ [ME]: { name: "Rin", role: "partner" } }),
            onInvite: (notice) => heard.push(notice),
        });
    }

    // Opening a session for the round dials the contact on the way up, and
    // that dial is the invitation. Dialling a second time put two notices on
    // the partner's screen for one share.
    it("puts one notice on their screen, not two", async () => {
        const heard: InviteNotice[] = [];
        await samMidRound(heard);

        await inviteContact(round, "sam");
        for (let i = 0; i < 20; i++) await Promise.resolve();

        expect(heard).toHaveLength(1);
        expect(heard[0].roundId).toBe(round.id);
    });

    it("still dials a contact onto a session that is already up", async () => {
        const heard: InviteNotice[] = [];
        await startForRound(round);
        await samMidRound(heard);

        await inviteContact(round, "sam");
        for (let i = 0; i < 20; i++) await Promise.resolve();

        expect(heard).toHaveLength(1);
    });
});

describe("opening another flow while a session is live", () => {
    // The replica is a singleton. A session left running for the round the
    // debater just left would be handed the new round's keystrokes, and would
    // merge the old partner's edits onto a grid they were never invited to.
    it("ends the session, even when the new round has nobody to dial", async () => {
        const shared = await startForRound(round);
        expect(shared).not.toBeNull();

        const priv = makeFlowRound({});
        expect(await resumeSession(priv)).toBeNull();

        expect(currentSession()).toBeNull();
        expect(useCollabStore.getState().status).toBe("off");
        expect(useCollabStore.getState().peers).toEqual([]);
    });

    it("keeps the session when the same round is opened again", async () => {
        const shared = await startForRound(round);
        rememberRoundPeers(round.id, ["sam"]);
        expect(await resumeSession(round)).toBe(shared);
        expect(currentSession()).toBe(shared);
    });

    it("stops pushing local edits into the round that was left", async () => {
        await startForRound(round);
        clearReplica();
        await resumeSession(makeFlowRound({}));
        // The bridge the replica pushes through is let go with the session, so
        // a write in the new round reaches nothing at all.
        expect(() => notifyLocalChange()).not.toThrow();
        expect(currentSession()).toBeNull();
    });
});

/**
 * The one route onto the network that no debater asks for: a flow that was
 * shared once carries its peers in its sidecar forever, and opening it again -
 * from Finder, from a file association, from a second launch - runs this.
 *
 * The first block is the positive control. An empty recorder only means the
 * gate held if the same call fills it when both switches are on, so those
 * assertions are what give the off cases their meaning.
 */
describe("resuming a round that was shared before", () => {
    beforeEach(async () => {
        // The beforeEach above ends any session, which binds the idle listener
        // again. Starting from nothing bound is what lets the recorder speak
        // for the resume alone.
        useFlowStore.setState({ collabListenEnabled: false });
        await syncInviteWatch();
        useFlowStore.setState({ collabListenEnabled: true });
        net.reset();
        rememberRoundPeers(round.id, ["sam"]);
    });

    it("binds an endpoint and dials the remembered peer, with both switches on", async () => {
        expect(await resumeSession(round)).not.toBeNull();
        expect(net.calls.filter((c) => c.op === "listen").map((c) => c.endpointId)).toEqual([ME]);
        expect(net.calls.filter((c) => c.op === "dial").map((c) => c.endpointId)).toEqual(["sam"]);
    });

    it("binds nothing and dials nobody while shared editing is off", async () => {
        useFlowStore.setState({ collabEnabled: false });
        expect(await resumeSession(round)).toBeNull();
        expect(net.calls).toEqual([]);
    });

    // A .ebb shared in October must not put this install back on the network
    // in March because it was double-clicked. The macOS local network prompt
    // belongs to the moment a debater asks to be reachable, never to startup.
    it("binds nothing and dials nobody while Listen for invites is off", async () => {
        useFlowStore.setState({ collabListenEnabled: false });
        expect(await resumeSession(round)).toBeNull();
        expect(net.calls).toEqual([]);
    });

    // The replica is a singleton, so the round being left has to lose its
    // session whatever the switches say - that is why the guard sits after it.
    it("still ends the session for the round being left while Listen for invites is off", async () => {
        expect(await startForRound(round)).not.toBeNull();
        useFlowStore.setState({ collabListenEnabled: false });

        expect(await resumeSession(makeFlowRound({}))).toBeNull();

        expect(currentSession()).toBeNull();
        expect(useCollabStore.getState().status).toBe("off");
    });
});
