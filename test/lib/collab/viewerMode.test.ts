/**
 * View-only mode, end to end over the memory transport.
 *
 * A viewer is the one peer that reads and does not write, so every claim here
 * is paired: the host refuses the viewer and accepts an editor in the same
 * shape, which is what makes a refusal mean the role held rather than the
 * message never landing.
 */

import { beforeEach, describe, expect, it } from "vitest";

import type { Contacts } from "@/lib/collab/contacts";
import { seedDoc } from "@/lib/collab/doc";
import { merge, type DroppedCell } from "@/lib/collab/merge";
import { applyOp, type OpContext } from "@/lib/collab/ops";
import { createMemoryNet } from "@/lib/collab/peerLinkMemory";
import { knownRoundPeers, knownRoundViewers, setRoundPeers } from "@/lib/collab/roundPeers";
import {
    startCollabSession,
    type CollabSession,
    type CollabSessionDeps,
} from "@/lib/collab/session";
import { createClock } from "@/lib/collab/stamp";
import { encodeTicket, parseTicket } from "@/lib/collab/ticket";
import type { CollabDoc, Role } from "@/lib/collab/types";
import { modelCol } from "@/lib/grid/colSpace";
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
    // What opening a round does before a session starts. The record is where a
    // read-only grant lives, so every test here starts with an empty one.
    setRoundPeers(shared.id, [], []);
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
async function hostAndGuest(role: Role, guestId = RAE) {
    const hostSide = replicaFor(shared, ALEX);
    const host = await open(ALEX, hostSide, { contacts: () => useFlowStore.getState().contacts });
    const ticket = encodeTicket(await host.share(role));

    const guestSide = replicaFor(shared, guestId);
    const guest = await open(guestId, guestSide, { ticket, role, dial: [ALEX] });
    await settle();
    return { host, guest, hostSide, guestSide, ticket };
}

describe("a view-only ticket", () => {
    /** Everything a ticket carries beside the role. */
    const stub = { endpointId: ALEX, roundId: "round_1", secret: "s".repeat(24), relay: true };

    it("carries the role, so the invite itself is what grants it", () => {
        expect(parseTicket(encodeTicket({ ...stub, role: "editor" }))!.role).toBe("editor");
        expect(parseTicket(encodeTicket({ ...stub, role: "viewer" }))!.role).toBe("viewer");
    });

    it("admits the holder as a viewer, and an editor ticket as an editor", async () => {
        const viewer = await hostAndGuest("viewer");
        expect(viewer.host.peers()[0].role).toBe("viewer");

        net.reset();
        clock.reset();
        const editor = await hostAndGuest("editor");
        expect(editor.host.peers()[0].role).toBe("editor");
    });

    it("tells the guest what it was admitted as, which is the only way it knows", async () => {
        const { host, guest } = await hostAndGuest("viewer");
        expect(guest.role()).toBe("viewer");
        // The host holds the file and is graded by nobody.
        expect(host.role()).toBe("editor");
    });

    it("leaves an editor an editor on both sides", async () => {
        const { host, guest } = await hostAndGuest("editor");
        expect(guest.role()).toBe("editor");
        expect(host.role()).toBe("editor");
    });

    it("cannot be spent as an editor by a guest that says it is one", async () => {
        const hostSide = replicaFor(shared, ALEX);
        const host = await open(ALEX, hostSide);
        const ticket = encodeTicket(await host.share("viewer"));

        // The guest asks for the role it wants. The host grants the role the
        // ticket names, and the ack says so to its face.
        const liarSide = replicaFor(shared, RAE);
        const liar = await open(RAE, liarSide, { ticket, role: "editor", dial: [ALEX] });
        await settle();

        expect(host.peers()[0].role).toBe("viewer");
        expect(liar.role()).toBe("viewer");
    });

    it("names the host an editor on the viewer's own peer list", async () => {
        // The chip reads this. A viewer whose list called the host view-only
        // would be reading its own role back at itself.
        const { guest } = await hostAndGuest("viewer");
        expect(guest.peers()).toHaveLength(1);
        expect(guest.peers()[0].endpointId).toBe(ALEX);
        expect(guest.peers()[0].role).toBe("editor");
    });
});

describe("what a viewer may do to the round", () => {
    it("reads the host's edits", async () => {
        const { host, hostSide, guestSide } = await hostAndGuest("viewer");
        hostSide.edit(sheetId, 0, 0, "host typed");
        host.notifyLocalChange();
        await settle();

        expect(guestSide.texts(sheetId)).toContain("host typed");
    });

    it("writes nothing back, while an editor's write in the same shape lands", async () => {
        const viewer = await hostAndGuest("viewer");
        viewer.guestSide.edit(sheetId, 0, 0, "viewer typed");
        viewer.guest.notifyLocalChange();
        await settle();
        expect(viewer.hostSide.texts(sheetId)).not.toContain("viewer typed");

        net.reset();
        clock.reset();
        const editor = await hostAndGuest("editor");
        editor.guestSide.edit(sheetId, 0, 0, "editor typed");
        editor.guest.notifyLocalChange();
        await settle();
        expect(editor.hostSide.texts(sheetId)).toContain("editor typed");
    });

    it("cannot reach an editor through the host, because the hub drops it first", async () => {
        const hostSide = replicaFor(shared, ALEX);
        const host = await open(ALEX, hostSide);

        const viewerSide = replicaFor(shared, RAE);
        const viewer = await open(RAE, viewerSide, {
            ticket: encodeTicket(await host.share("viewer")),
            role: "viewer",
            dial: [ALEX],
        });
        const editorSide = replicaFor(shared, SAM);
        await open(SAM, editorSide, {
            ticket: encodeTicket(await host.share("editor")),
            role: "editor",
            dial: [ALEX],
        });
        await settle();
        expect(host.peers()).toHaveLength(2);

        viewerSide.edit(sheetId, 0, 0, "viewer typed");
        viewer.notifyLocalChange();
        await settle();

        // The star has one hub. Nothing the hub refused is forwarded, so the
        // editor never sees it either.
        expect(hostSide.texts(sheetId)).not.toContain("viewer typed");
        expect(editorSide.texts(sheetId)).not.toContain("viewer typed");
    });

    it("still shows the host where it is looking, marked as the read-only one", async () => {
        // Presence is not a write. A viewer following the round is the whole
        // reason the role exists, and the mark is what lets the grid hide a
        // cursor that can never refuse a keystroke.
        const { host, guest } = await hostAndGuest("viewer");
        guest.setCursor({ sheetId, col: modelCol(1), row: 3 });
        await settle();

        expect(getPresences()[0]).toMatchObject({
            endpointId: RAE,
            col: 1,
            row: 3,
            editing: false,
            readOnly: true,
        });
        await host.stop();
    });

    it("cannot claim the cell the host is typing in, while an editor can", async () => {
        // A claim refuses the local debater's keystroke, so it is a write under
        // another name. A viewer that could make one would lock the host out of
        // the cell they are mid-speech in.
        const viewer = await hostAndGuest("viewer");
        viewer.guest.setPresence({ sheetId, col: modelCol(1), row: 3 });
        await settle();
        expect(getPresences()[0]).toMatchObject({ endpointId: RAE, col: 1, row: 3 });
        expect(getPresences()[0].editing).toBe(false);
        await viewer.host.stop();
        await viewer.guest.stop();

        net.reset();
        clock.reset();
        setPresences([]);
        const editor = await hostAndGuest("editor", SAM);
        editor.guest.setPresence({ sheetId, col: modelCol(1), row: 3 });
        await settle();
        expect(getPresences()[0]).toMatchObject({
            endpointId: SAM,
            editing: true,
            readOnly: false,
        });
        await editor.host.stop();
        await editor.guest.stop();
    });
});

describe("a viewer that comes back", () => {
    /** What the next open of the round is: its record, and two fresh sessions. */
    async function reopen(peers: readonly string[], readOnly: readonly string[]) {
        net.reset();
        clock.reset();
        setPresences([]);
        setRoundPeers(shared.id, peers, readOnly);
        const hostSide = replicaFor(shared, ALEX);
        const host = await open(ALEX, hostSide, {
            dial: knownRoundPeers(shared.id),
            contacts: () => useFlowStore.getState().contacts,
        });
        const raeSide = replicaFor(shared, RAE);
        const rae = await open(RAE, raeSide, { dial: [ALEX] });
        await settle();
        return { host, rae, hostSide, raeSide };
    }

    /** Rae typing into the round and pushing it, which is the write under test. */
    async function raeWrites(rae: CollabSession, raeSide: Replica): Promise<void> {
        raeSide.edit(sheetId, 0, 0, "rae typed");
        rae.notifyLocalChange();
        await settle();
    }

    it("is still a viewer on the session that admitted it", async () => {
        const { host, guest } = await hostAndGuest("viewer");
        await guest.stop();
        await settle();
        expect(host.peers()).toEqual([]);

        // The round remembers who to redial and the session remembers what
        // they were. No second ticket is spent.
        const again = await open(RAE, replicaFor(shared, RAE), { dial: [ALEX] });
        await settle();
        expect(host.peers()[0].role).toBe("viewer");
        expect(again.role()).toBe("viewer");
    });

    it("comes back an editor when the round never marked them read-only", async () => {
        // Membership with no mark beside it is an editor, which is the common
        // case: a round remembers everyone it was shared with, and only a
        // read-only grant is a restriction worth carrying. A saved contact
        // beside it changes nothing, because a contact grades nobody.
        useFlowStore.setState({ contacts: { [RAE]: { name: "Rae" } } });
        const { host, rae, hostSide, raeSide } = await reopen([RAE], []);

        expect(host.peers()[0].role).toBe("editor");
        await raeWrites(rae, raeSide);
        expect(hostSide.texts(sheetId)).toContain("rae typed");

        await host.stop();
        await rae.stop();
    });

    // A contact is a name and an address. It says who somebody is and never
    // what they may do, so a table that holds them, a table that calls them
    // something else, and no table at all are one answer: the round's mark.
    it("is still a viewer whatever the contact table holds", async () => {
        const first = await hostAndGuest("viewer");
        expect(knownRoundViewers(shared.id)).toEqual([RAE]);
        await first.guest.stop();
        await first.host.stop();

        const tables: Contacts[] = [{}, { [RAE]: { name: "Rae" } }];
        for (const contacts of tables) {
            useFlowStore.setState({ contacts });
            const { host, rae, hostSide, raeSide } = await reopen([RAE], [RAE]);

            expect(host.peers()[0].role).toBe("viewer");
            expect(rae.role()).toBe("viewer");
            await raeWrites(rae, raeSide);
            expect(hostSide.texts(sheetId)).not.toContain("rae typed");

            await host.stop();
            await rae.stop();
        }
    });

    // The contact table used to hold the only record of the restriction while
    // the round held the record of the membership, so a debater who removed the
    // contact - the gesture that most looks like withdrawing trust - promoted
    // the peer they meant to keep read-only.
    it("is still a viewer after the debater removes their contact", async () => {
        useFlowStore.setState({ contacts: { [RAE]: { name: "Rae" } } });
        const first = await hostAndGuest("viewer");
        expect(first.host.peers()[0].role).toBe("viewer");
        expect(knownRoundViewers(shared.id)).toEqual([RAE]);
        await first.guest.stop();
        await first.host.stop();

        useFlowStore.setState({ contacts: {} });
        const { host, rae, hostSide, raeSide } = await reopen([RAE], [RAE]);

        expect(host.peers()[0].role).toBe("viewer");
        expect(rae.role()).toBe("viewer");
        await raeWrites(rae, raeSide);
        expect(hostSide.texts(sheetId)).not.toContain("rae typed");

        await host.stop();
        await rae.stop();
    });

    // The grant a session made only lasted as long as the session, so a peer
    // handed the wider role mid-round was read-only again the next morning:
    // the durable record was a contact row saved while they were still a
    // viewer. An invitation is the debater deciding, and it moves the mark.
    it("stays an editor after a promotion, on the next open of the round", async () => {
        useFlowStore.setState({ contacts: { [RAE]: { name: "Rae" } } });
        const first = await hostAndGuest("viewer");
        expect(first.host.peers()[0].role).toBe("viewer");
        expect(knownRoundViewers(shared.id)).toEqual([RAE]);

        await first.host.invite(RAE, "editor");
        await settle();
        expect(first.host.peers()[0].role).toBe("editor");
        // The mark moved with the grant, so the sidecar carries the promotion.
        const readOnly = knownRoundViewers(shared.id);
        expect(readOnly).toEqual([]);
        await first.guest.stop();
        await first.host.stop();

        const { host, rae, hostSide, raeSide } = await reopen([RAE], readOnly);
        expect(host.peers()[0].role).toBe("editor");
        await raeWrites(rae, raeSide);
        expect(hostSide.texts(sheetId)).toContain("rae typed");

        await host.stop();
        await rae.stop();
    });
});

describe("a viewer on a link the host dialled", () => {
    /** A host that reopened a round it shared read-only, and dials the viewer. */
    async function hostDialsViewer() {
        setRoundPeers(shared.id, [RAE], [RAE]);
        // The viewer comes up first and cannot reach a host that is not
        // listening yet, so the link that lands is the host's own dial.
        const viewerSide = replicaFor(shared, RAE);
        const viewer = await open(RAE, viewerSide, { dial: [ALEX] });
        const hostSide = replicaFor(shared, ALEX);
        const host = await open(ALEX, hostSide, {
            dial: knownRoundPeers(shared.id),
            contacts: () => useFlowStore.getState().contacts,
        });
        await settle();
        return { host, viewer, hostSide, viewerSide };
    }

    // The comment on the dial path assumed the dialler is always a guest, and
    // hardcoded read-only off. The host dials too - every remembered peer when
    // it reopens a round, and a contact on invite - so on those links the host
    // applied the document of a peer it had granted read only.
    it("has its writes dropped, the same as on a link the host answered", async () => {
        const { host, viewer, hostSide, viewerSide } = await hostDialsViewer();
        expect(host.peers()).toHaveLength(1);
        expect(host.peers()[0].role).toBe("viewer");

        viewerSide.edit(sheetId, 0, 0, "viewer typed");
        viewer.notifyLocalChange();
        await settle();
        expect(hostSide.texts(sheetId)).not.toContain("viewer typed");

        await host.stop();
        await viewer.stop();
    });

    it("cannot claim a cell on that link either", async () => {
        const { host, viewer } = await hostDialsViewer();
        viewer.setPresence({ sheetId, col: modelCol(2), row: 4 });
        await settle();

        expect(getPresences()[0]).toMatchObject({ endpointId: RAE, col: 2, row: 4 });
        expect(getPresences()[0].editing).toBe(false);
        expect(getPresences()[0].readOnly).toBe(true);

        await host.stop();
        await viewer.stop();
    });

    it("still reads the host's edits, which is the whole point of the role", async () => {
        const { host, viewer, hostSide, viewerSide } = await hostDialsViewer();
        hostSide.edit(sheetId, 0, 0, "host typed");
        host.notifyLocalChange();
        await settle();
        expect(viewerSide.texts(sheetId)).toContain("host typed");

        await host.stop();
        await viewer.stop();
    });
});
