import { beforeEach, describe, expect, it } from "vitest";

import { seedDoc } from "@/lib/collab/doc";
import { merge } from "@/lib/collab/merge";
import { applyOp, type OpContext } from "@/lib/collab/ops";
import { createMemoryNet, memoryRelay } from "@/lib/collab/peerLinkMemory";
import { HEARTBEAT_MS } from "@/lib/collab/presence";
import { knownRoundRelays } from "@/lib/collab/roundPeers";
import { startCollabSession, type CollabSession } from "@/lib/collab/session";
import { createClock } from "@/lib/collab/stamp";
import { encodeTicket } from "@/lib/collab/ticket";
import type { CollabDoc } from "@/lib/collab/types";
import { modelCol } from "@/lib/grid/colSpace";
import { getPresences, setPresences } from "@/lib/grid/presenceBridge";
import { makeFlowRound, type FlowRound } from "@/lib/model/flow";
import { useFlowStore } from "@/lib/store/useFlowStore";

const net = createMemoryNet();

/** What iroh hands back. A ticket names the host, so the host holds a real one. */
const ALEX = "a".repeat(64);

/** The scheduler both sessions run on, so the test owns time. */
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
        reset() {
            pending = [];
            now = 0;
        },
    };
}

const clock = manualClock();

/**
 * Settles the handshake and one coalesced push. The push debounce is 30ms on
 * the injected clock; the repair tick is seconds away and never fires here.
 */
async function settle(): Promise<void> {
    for (let i = 0; i < 10; i++) await Promise.resolve();
    clock.advance(50);
    for (let i = 0; i < 10; i++) await Promise.resolve();
}

function round(): FlowRound {
    const r = makeFlowRound({});
    const flow = r.sheets.find((s) => s.kind !== "cx")!;
    flow.data = [
        ["perm", "link"],
        ["cap bad", "turn"],
    ];
    return r;
}

/** One peer's own replica plus the session that syncs it. */
function replicaFor(base: FlowRound, actor: string) {
    let doc = seedDoc(base);
    let t = actor === ALEX ? 1_000 : 5_000;
    const ctx: OpContext = { actor, clock: createClock(actor, () => t++) };
    return {
        doc: () => doc,
        apply: (incoming: CollabDoc) => {
            const result = merge(doc, incoming);
            doc = result.doc;
            return result.dropped;
        },
        edit(sheetId: string, col: number, row: number, text: string) {
            doc = applyOp(doc, { kind: "cellText", sheetId, col, row, text }, ctx);
        },
    };
}

let shared: FlowRound;
let sheetId: string;

beforeEach(() => {
    net.reset();
    clock.reset();
    setPresences([]);
    useFlowStore.setState({ collabEnabled: true, collabRelayEnabled: true });
    shared = round();
    sheetId = shared.sheets.find((s) => s.kind !== "cx")!.id;
});

async function hostAndGuest(): Promise<{
    host: CollabSession;
    guest: CollabSession;
    hostSide: ReturnType<typeof replicaFor>;
    guestSide: ReturnType<typeof replicaFor>;
    ticket: string;
}> {
    const hostSide = replicaFor(shared, ALEX);
    const host = (await startCollabSession({
        createLink: net.create(ALEX),
        roundId: shared.id,
        appVersion: "0.11.0",
        doc: hostSide.doc,
        apply: hostSide.apply,
        schedule: clock.schedule,
    }))!;
    const ticket = encodeTicket(await host.share("partner"));

    const guestSide = replicaFor(shared, "sam");
    const guest = (await startCollabSession({
        createLink: net.create("sam"),
        roundId: shared.id,
        appVersion: "0.11.0",
        doc: guestSide.doc,
        apply: guestSide.apply,
        ticket,
        dial: [ALEX],
        schedule: clock.schedule,
    }))!;
    await settle();
    return { host, guest, hostSide, guestSide, ticket };
}

describe("a hosted session", () => {
    it("admits a guest that presents the ticket", async () => {
        const { host, guest } = await hostAndGuest();
        expect(host.peers().map((p) => p.endpointId)).toEqual(["sam"]);
        expect(guest.peers().map((p) => p.endpointId)).toEqual([ALEX]);
        expect(host.peers()[0].role).toBe("partner");
    });

    it("reports the connection type each peer got", async () => {
        const { host } = await hostAndGuest();
        expect(["direct", "relayed"]).toContain(host.peers()[0].connectionType);
    });

    /**
     * A guest two networks apart has an EndpointId and nothing to send a
     * packet to: mDNS answers across a room and no further. The ticket is the
     * only thing that has been handed over by then, so it is what carries the
     * host's relay, and the dial that redeems it looks there.
     */
    it("mints a ticket naming where the host is, and dials there", async () => {
        await hostAndGuest();
        const dials = net.calls.filter((c) => c.op === "dial");
        expect(dials).toEqual([{ op: "dial", endpointId: ALEX, relayUrl: memoryRelay(ALEX) }]);
    });

    it("names no relay in the ticket when the host would not use one", async () => {
        useFlowStore.setState({ collabRelayEnabled: false });
        const hostSide = replicaFor(shared, ALEX);
        const host = (await startCollabSession({
            createLink: net.create(ALEX),
            roundId: shared.id,
            appVersion: "0.11.0",
            doc: hostSide.doc,
            apply: hostSide.apply,
            schedule: clock.schedule,
        }))!;
        expect((await host.share("partner")).relayUrl).toBeUndefined();
    });

    /**
     * The ticket is spent once. Every reconnect after it dials by EndpointId,
     * so where the peer was found has to outlive the connection that found it
     * or a link that blips two networks apart never comes back.
     */
    it("remembers where each peer was found, on both sides", async () => {
        const { host, guest } = await hostAndGuest();
        expect(host.peers()[0].relayUrl).toBe(memoryRelay("sam"));
        expect(guest.peers()[0].relayUrl).toBe(memoryRelay(ALEX));
        expect(knownRoundRelays(shared.id)).toEqual({
            sam: memoryRelay("sam"),
            [ALEX]: memoryRelay(ALEX),
        });
    });

    it("refuses an unknown peer with no ticket, and shows nothing", async () => {
        const hostSide = replicaFor(shared, ALEX);
        const host = (await startCollabSession({
            createLink: net.create(ALEX),
            roundId: shared.id,
            appVersion: "0.11.0",
            doc: hostSide.doc,
            apply: hostSide.apply,
            schedule: clock.schedule,
        }))!;

        const strangerSide = replicaFor(shared, "mallory");
        await startCollabSession({
            createLink: net.create("mallory"),
            roundId: shared.id,
            appVersion: "0.11.0",
            doc: strangerSide.doc,
            apply: strangerSide.apply,
            dial: [ALEX],
            schedule: clock.schedule,
        });
        await settle();
        expect(host.peers()).toEqual([]);
    });

    it("spends the ticket, so a second stranger cannot reuse it", async () => {
        const { host, ticket } = await hostAndGuest();
        const strangerSide = replicaFor(shared, "mallory");
        await startCollabSession({
            createLink: net.create("mallory"),
            roundId: shared.id,
            appVersion: "0.11.0",
            doc: strangerSide.doc,
            apply: strangerSide.apply,
            ticket,
            dial: [ALEX],
            schedule: clock.schedule,
        });
        await settle();
        expect(host.peers().map((p) => p.endpointId)).toEqual(["sam"]);
    });

    it("re-admits a known peer with no ticket at all", async () => {
        const { host, guest } = await hostAndGuest();
        await guest.stop();
        await settle();
        expect(host.peers()).toEqual([]);

        const againSide = replicaFor(shared, "sam");
        const again = (await startCollabSession({
            createLink: net.create("sam"),
            roundId: shared.id,
            appVersion: "0.11.0",
            doc: againSide.doc,
            apply: againSide.apply,
            dial: [ALEX],
            schedule: clock.schedule,
        }))!;
        await settle();
        expect(host.peers().map((p) => p.endpointId)).toEqual(["sam"]);
        await again.stop();
    });
});

describe("the name each side broadcasts", () => {
    async function named(hostName?: string, guestName?: string) {
        const hostSide = replicaFor(shared, ALEX);
        const host = (await startCollabSession({
            createLink: net.create(ALEX),
            roundId: shared.id,
            appVersion: "0.11.0",
            doc: hostSide.doc,
            apply: hostSide.apply,
            displayName: hostName,
            schedule: clock.schedule,
        }))!;
        const ticket = encodeTicket(await host.share("partner"));

        const guestSide = replicaFor(shared, "sam");
        const guest = (await startCollabSession({
            createLink: net.create("sam"),
            roundId: shared.id,
            appVersion: "0.11.0",
            doc: guestSide.doc,
            apply: guestSide.apply,
            displayName: guestName,
            ticket,
            dial: [ALEX],
            schedule: clock.schedule,
        }))!;
        await settle();
        return { host, guest };
    }

    it("reaches the host from the guest's hello", async () => {
        const { host } = await named("Alex", "Rin");
        expect(host.peers()[0].name).toBe("Rin");
    });

    it("comes back to the guest on the ack, so naming works both ways", async () => {
        const { guest } = await named("Alex", "Rin");
        expect(guest.peers()[0].name).toBe("Alex");
    });

    it("leaves a peer nameless when the far side broadcasts nothing", async () => {
        const { host, guest } = await named(undefined, undefined);
        expect(host.peers()[0].name).toBeUndefined();
        expect(guest.peers()[0].name).toBeUndefined();
    });
});

describe("editing across a session", () => {
    it("carries an edit from the host to the guest", async () => {
        const { host, hostSide, guestSide } = await hostAndGuest();
        hostSide.edit(sheetId, 0, 0, "host typed");
        host.notifyLocalChange();
        await settle();

        const cells = Object.values(guestSide.doc().sheets[sheetId].cells);
        expect(cells.map((c) => c.text)).toContain("host typed");
    });

    it("carries an edit from the guest to the host", async () => {
        const { guest, hostSide, guestSide } = await hostAndGuest();
        guestSide.edit(sheetId, 1, 1, "guest typed");
        guest.notifyLocalChange();
        await settle();

        const cells = Object.values(hostSide.doc().sheets[sheetId].cells);
        expect(cells.map((c) => c.text)).toContain("guest typed");
    });

    it("converges when both type at once", async () => {
        const { host, guest, hostSide, guestSide } = await hostAndGuest();
        hostSide.edit(sheetId, 0, 0, "from alex");
        guestSide.edit(sheetId, 1, 0, "from sam");
        host.notifyLocalChange();
        guest.notifyLocalChange();
        await settle();

        expect(hostSide.doc()).toEqual(guestSide.doc());
    });

    it("drops a coach's writes, because the host enforces the role", async () => {
        const hostSide = replicaFor(shared, ALEX);
        const host = (await startCollabSession({
            createLink: net.create(ALEX),
            roundId: shared.id,
            appVersion: "0.11.0",
            doc: hostSide.doc,
            apply: hostSide.apply,
            schedule: clock.schedule,
        }))!;
        const ticket = encodeTicket(await host.share("coach"));

        const coachSide = replicaFor(shared, "coach");
        const coach = (await startCollabSession({
            createLink: net.create("coach"),
            roundId: shared.id,
            appVersion: "0.11.0",
            doc: coachSide.doc,
            apply: coachSide.apply,
            ticket,
            role: "coach",
            dial: [ALEX],
            schedule: clock.schedule,
        }))!;
        await settle();
        expect(host.peers()[0].role).toBe("coach");

        coachSide.edit(sheetId, 0, 0, "coach typed");
        coach.notifyLocalChange();
        await settle();

        const texts = Object.values(hostSide.doc().sheets[sheetId].cells).map((c) => c.text);
        expect(texts).not.toContain("coach typed");
    });
});

describe("the opt-in gate still holds", () => {
    it("hands back nothing with shared editing off", async () => {
        useFlowStore.setState({ collabEnabled: false });
        const side = replicaFor(shared, ALEX);
        const session = await startCollabSession({
            createLink: net.create(ALEX),
            roundId: shared.id,
            appVersion: "0.11.0",
            doc: side.doc,
            apply: side.apply,
            schedule: clock.schedule,
        });
        expect(session).toBeNull();
        expect(net.calls).toEqual([]);
    });
});

describe("presence across a session", () => {
    it("shows the cell a partner has an editor open on", async () => {
        const { host, guest } = await hostAndGuest();
        guest.setPresence({ sheetId, col: modelCol(1), row: 4 });
        await settle();
        expect(getPresences()).toHaveLength(1);
        expect(getPresences()[0]).toMatchObject({
            endpointId: "sam",
            col: 1,
            row: 4,
            editing: true,
        });
        await host.stop();
    });

    it("releases it the moment their editor closes", async () => {
        const { guest } = await hostAndGuest();
        guest.setPresence({ sheetId, col: modelCol(1), row: 4 });
        await settle();
        guest.setPresence(null);
        await settle();
        expect(getPresences()).toEqual([]);
    });

    it("leaves an unreachable peer holding nothing", async () => {
        const { guest } = await hostAndGuest();
        guest.setPresence({ sheetId, col: modelCol(0), row: 0 });
        await settle();
        expect(getPresences()).toHaveLength(1);

        // The link drops. Nothing waits for the heartbeat to lapse.
        await guest.stop();
        await settle();
        expect(getPresences()).toEqual([]);
    });
});

describe("a partner's cursor across a session", () => {
    it("shows the cell they are on, claiming nothing", async () => {
        const { host, guest } = await hostAndGuest();
        guest.setCursor({ sheetId, col: modelCol(1), row: 4 });
        await settle();
        expect(getPresences()).toHaveLength(1);
        expect(getPresences()[0]).toMatchObject({
            endpointId: "sam",
            col: 1,
            row: 4,
            editing: false,
        });
        await host.stop();
    });

    it("keeps arrowing off the wire, and the heartbeat carries where they landed", async () => {
        const { host, guest } = await hostAndGuest();
        guest.setCursor({ sheetId, col: modelCol(0), row: 0 });
        await settle();

        // A debater holding the down arrow moves faster than anything needs to
        // hear about, so only the first of a burst goes out at once.
        for (let row = 1; row <= 8; row++) guest.setCursor({ sheetId, col: modelCol(0), row });
        await settle();
        expect(getPresences()[0]).toMatchObject({ row: 0 });

        clock.advance(HEARTBEAT_MS);
        for (let i = 0; i < 10; i++) await Promise.resolve();
        expect(getPresences()[0]).toMatchObject({ row: 8, editing: false });
        await host.stop();
    });

    it("hands the cell back the moment they leave the grid", async () => {
        const { host, guest } = await hostAndGuest();
        guest.setCursor({ sheetId, col: modelCol(2), row: 1 });
        await settle();
        expect(getPresences()).toHaveLength(1);

        guest.setCursor(null);
        await settle();
        expect(getPresences()).toEqual([]);
        await host.stop();
    });

    it("does not downgrade an open editor, because the editor speaks for both", async () => {
        const { host, guest } = await hostAndGuest();
        guest.setCursor({ sheetId, col: modelCol(1), row: 3 });
        guest.setPresence({ sheetId, col: modelCol(1), row: 3 });
        await settle();
        expect(getPresences()[0].editing).toBe(true);

        // A selection cannot move under an open editor, but a stray claim must
        // not unlock the cell either.
        guest.setCursor({ sheetId, col: modelCol(1), row: 3 });
        clock.advance(HEARTBEAT_MS * 2);
        for (let i = 0; i < 10; i++) await Promise.resolve();
        expect(getPresences()[0].editing).toBe(true);
        await host.stop();
    });

    it("goes back to a plain cursor when the editor closes on the cell they keep", async () => {
        const { host, guest } = await hostAndGuest();
        guest.setCursor({ sheetId, col: modelCol(1), row: 3 });
        guest.setPresence({ sheetId, col: modelCol(1), row: 3 });
        await settle();
        expect(getPresences()[0].editing).toBe(true);

        // No blank frame: the release carries the cursor with it, so the cell
        // never stops showing where the partner is.
        guest.setPresence(null);
        await settle();
        expect(getPresences()).toHaveLength(1);
        expect(getPresences()[0]).toMatchObject({ col: 1, row: 3, editing: false });
        await host.stop();
    });

    it("stops the heartbeat once a peer is nowhere, so an idle pane costs no timer", async () => {
        const { host, guest } = await hostAndGuest();
        guest.setCursor({ sheetId, col: modelCol(0), row: 0 });
        await settle();
        guest.setCursor(null);
        await settle();

        // Nothing left to say: advancing far past the TTL puts no message on
        // the wire and leaves the table empty rather than resurrecting a cell.
        clock.advance(HEARTBEAT_MS * 20);
        for (let i = 0; i < 10; i++) await Promise.resolve();
        expect(getPresences()).toEqual([]);
        await host.stop();
    });
});
