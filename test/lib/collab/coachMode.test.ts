/**
 * View-only mode, end to end over the memory transport.
 *
 * A coach is the one peer that reads and does not write, so every claim here
 * is paired: the host refuses the coach and accepts a partner in the same
 * shape, which is what makes a refusal mean the role held rather than the
 * message never landing.
 */

import { beforeEach, describe, expect, it } from "vitest";

import type { Contacts } from "@/lib/collab/contacts";
import { seedDoc } from "@/lib/collab/doc";
import { merge, type DroppedCell } from "@/lib/collab/merge";
import { applyOp, type OpContext } from "@/lib/collab/ops";
import { createMemoryNet } from "@/lib/collab/peerLinkMemory";
import {
    startCollabSession,
    type CollabSession,
    type CollabSessionDeps,
} from "@/lib/collab/session";
import { createClock } from "@/lib/collab/stamp";
import { encodeTicket, parseTicket } from "@/lib/collab/ticket";
import type { CollabDoc } from "@/lib/collab/types";
import { getPresences, setPresences } from "@/lib/grid/presenceBridge";
import { makeFlowRound, type FlowRound } from "@/lib/model/flow";
import { useFlowStore } from "@/lib/store/useFlowStore";

const net = createMemoryNet();

/** What iroh hands back. A ticket names the host, so the host holds a real one. */
const ALEX = "a".repeat(64);
const RAE = "r".repeat(64);
const SAM = "s".repeat(64);

/** The scheduler every session runs on, so the test owns time. */
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

/** Settles the handshake and one coalesced push, on the injected clock. */
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

/** One peer's own replica plus the writes it makes into it. */
interface Replica {
    doc(): CollabDoc;
    apply(incoming: CollabDoc): DroppedCell[];
    edit(sheetId: string, col: number, row: number, text: string): void;
    /** Every cell's text on one sheet. A cell emptied by a peer carries null. */
    texts(sheetId: string): (string | null)[];
}

function replicaFor(base: FlowRound, actor: string): Replica {
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
        texts(sheetId: string) {
            return Object.values(doc.sheets[sheetId].cells).map((c) => c.text);
        },
    };
}

let shared: FlowRound;
let sheetId: string;

beforeEach(() => {
    net.reset();
    clock.reset();
    setPresences([]);
    useFlowStore.setState({ collabEnabled: true, collabRelayEnabled: true, contacts: {} });
    shared = round();
    sheetId = shared.sheets.find((s) => s.kind !== "cx")!.id;
});

async function open(
    actor: string,
    side: Replica,
    over: Partial<CollabSessionDeps> = {},
): Promise<CollabSession> {
    const session = await startCollabSession({
        createLink: net.create(actor),
        roundId: shared.id,
        appVersion: "0.11.0",
        doc: side.doc,
        apply: side.apply,
        schedule: clock.schedule,
        ...over,
    });
    return session!;
}

/** A host with a guest that redeemed a ticket of the given role. */
async function hostAndGuest(role: "partner" | "coach", guestId = RAE) {
    const hostSide = replicaFor(shared, ALEX);
    const host = await open(ALEX, hostSide, { contacts: () => useFlowStore.getState().contacts });
    const ticket = encodeTicket(host.share(role));

    const guestSide = replicaFor(shared, guestId);
    const guest = await open(guestId, guestSide, { ticket, role, dial: [ALEX] });
    await settle();
    return { host, guest, hostSide, guestSide, ticket };
}

describe("a view-only ticket", () => {
    /** Everything a ticket carries beside the role. */
    const stub = { endpointId: ALEX, roundId: "round_1", secret: "s".repeat(24), relay: true };

    it("carries the role, so the invite itself is what grants it", () => {
        expect(parseTicket(encodeTicket({ ...stub, role: "partner" }))!.role).toBe("partner");
        expect(parseTicket(encodeTicket({ ...stub, role: "coach" }))!.role).toBe("coach");
    });

    it("admits the holder as a coach, and a partner ticket as a partner", async () => {
        const viewer = await hostAndGuest("coach");
        expect(viewer.host.peers()[0].role).toBe("coach");

        net.reset();
        clock.reset();
        const editor = await hostAndGuest("partner");
        expect(editor.host.peers()[0].role).toBe("partner");
    });

    it("tells the guest what it was admitted as, which is the only way it knows", async () => {
        const { host, guest } = await hostAndGuest("coach");
        expect(guest.role()).toBe("coach");
        // The host holds the file and is graded by nobody.
        expect(host.role()).toBe("partner");
    });

    it("leaves a partner a partner on both sides", async () => {
        const { host, guest } = await hostAndGuest("partner");
        expect(guest.role()).toBe("partner");
        expect(host.role()).toBe("partner");
    });

    it("cannot be spent as a partner by a guest that says it is one", async () => {
        const hostSide = replicaFor(shared, ALEX);
        const host = await open(ALEX, hostSide);
        const ticket = encodeTicket(host.share("coach"));

        // The guest asks for the role it wants. The host grants the role the
        // ticket names, and the ack says so to its face.
        const liarSide = replicaFor(shared, RAE);
        const liar = await open(RAE, liarSide, { ticket, role: "partner", dial: [ALEX] });
        await settle();

        expect(host.peers()[0].role).toBe("coach");
        expect(liar.role()).toBe("coach");
    });

    it("names the host a partner on the coach's own peer list", async () => {
        // The chip reads this. A coach whose list called the host view-only
        // would be reading its own role back at itself.
        const { guest } = await hostAndGuest("coach");
        expect(guest.peers()).toHaveLength(1);
        expect(guest.peers()[0].endpointId).toBe(ALEX);
        expect(guest.peers()[0].role).toBe("partner");
    });
});

describe("what a coach may do to the round", () => {
    it("reads the host's edits", async () => {
        const { host, hostSide, guestSide } = await hostAndGuest("coach");
        hostSide.edit(sheetId, 0, 0, "host typed");
        host.notifyLocalChange();
        await settle();

        expect(guestSide.texts(sheetId)).toContain("host typed");
    });

    it("writes nothing back, while a partner's write in the same shape lands", async () => {
        const viewer = await hostAndGuest("coach");
        viewer.guestSide.edit(sheetId, 0, 0, "coach typed");
        viewer.guest.notifyLocalChange();
        await settle();
        expect(viewer.hostSide.texts(sheetId)).not.toContain("coach typed");

        net.reset();
        clock.reset();
        const editor = await hostAndGuest("partner");
        editor.guestSide.edit(sheetId, 0, 0, "partner typed");
        editor.guest.notifyLocalChange();
        await settle();
        expect(editor.hostSide.texts(sheetId)).toContain("partner typed");
    });

    it("cannot reach a partner through the host, because the hub drops it first", async () => {
        const hostSide = replicaFor(shared, ALEX);
        const host = await open(ALEX, hostSide);

        const coachSide = replicaFor(shared, RAE);
        const coach = await open(RAE, coachSide, {
            ticket: encodeTicket(host.share("coach")),
            role: "coach",
            dial: [ALEX],
        });
        const partnerSide = replicaFor(shared, SAM);
        await open(SAM, partnerSide, {
            ticket: encodeTicket(host.share("partner")),
            role: "partner",
            dial: [ALEX],
        });
        await settle();
        expect(host.peers()).toHaveLength(2);

        coachSide.edit(sheetId, 0, 0, "coach typed");
        coach.notifyLocalChange();
        await settle();

        // The star has one hub. Nothing the hub refused is forwarded, so the
        // partner never sees it either.
        expect(hostSide.texts(sheetId)).not.toContain("coach typed");
        expect(partnerSide.texts(sheetId)).not.toContain("coach typed");
    });

    it("still shows the host where it is looking", async () => {
        // Presence is not a write. A coach following the round is the whole
        // reason the role exists.
        const { host, guest } = await hostAndGuest("coach");
        guest.setCursor({ sheetId, col: 1, row: 3 });
        await settle();

        expect(getPresences()[0]).toMatchObject({
            endpointId: RAE,
            col: 1,
            row: 3,
            editing: false,
        });
        await host.stop();
    });
});

describe("a coach that comes back", () => {
    it("is still a coach on the session that admitted it", async () => {
        const { host, guest } = await hostAndGuest("coach");
        await guest.stop();
        await settle();
        expect(host.peers()).toEqual([]);

        // The round remembers who to redial and the session remembers what
        // they were. No second ticket is spent.
        const again = await open(RAE, replicaFor(shared, RAE), { dial: [ALEX] });
        await settle();
        expect(host.peers()[0].role).toBe("coach");
        expect(again.role()).toBe("coach");
    });

    it("is still a coach after the host reopens the round, once they are a saved contact", async () => {
        // The round file remembers who to redial but not what they were, so
        // the contact table is where a role outlives a session. This is the
        // difference between a coach and a partner surviving an app restart.
        const contacts: Contacts = { [RAE]: { name: "Coach", role: "coach" } };
        useFlowStore.setState({ contacts });

        const hostSide = replicaFor(shared, ALEX);
        const host = await open(ALEX, hostSide, {
            dial: [RAE],
            contacts: () => useFlowStore.getState().contacts,
        });
        // No ticket: a peer the round already knows is admitted by EndpointId.
        const coachSide = replicaFor(shared, RAE);
        const coach = await open(RAE, coachSide, { dial: [ALEX] });
        await settle();

        expect(host.peers()[0].role).toBe("coach");
        expect(coach.role()).toBe("coach");

        coachSide.edit(sheetId, 0, 0, "coach typed");
        coach.notifyLocalChange();
        await settle();
        expect(hostSide.texts(sheetId)).not.toContain("coach typed");
    });

    it("comes back a partner when nobody saved them, which is what the round remembers", async () => {
        // Documented, not solved: a peer the round knows and the contact table
        // does not is a partner. Read-only is a fat-finger guard, and the way
        // to make it outlive a session is to save the contact.
        const host = await open(ALEX, replicaFor(shared, ALEX), {
            dial: [RAE],
            contacts: () => useFlowStore.getState().contacts,
        });
        await open(RAE, replicaFor(shared, RAE), { dial: [ALEX] });
        await settle();
        expect(host.peers()[0].role).toBe("partner");
    });
});
