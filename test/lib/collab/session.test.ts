import { beforeEach, describe, expect, it } from "vitest";

import { seedDoc } from "@/lib/collab/doc";
import { VERSION_SKEW } from "@/lib/collab/handshake";
import { merge } from "@/lib/collab/merge";
import {
    PROTOCOL_MAJOR,
    type PeerConn,
    type PeerLinkFactory,
    type WireMessage,
} from "@/lib/collab/peerLink";
import { HANDSHAKE_MS } from "@/lib/collab/peerLink";
import { createMemoryNet } from "@/lib/collab/peerLinkMemory";
import { forgetRoundPeers, knownRoundPeers, setRoundPeers } from "@/lib/collab/roundPeers";
import { startCollabSession, type CollabPeer, type CollabSession } from "@/lib/collab/session";
import { encodeTicket } from "@/lib/collab/ticket";
import type { CollabDoc } from "@/lib/collab/types";
import { getPresences } from "@/lib/grid/presenceBridge";
import { makeFlowRound, type FlowRound } from "@/lib/model/flow";
import { useFlowStore } from "@/lib/store/useFlowStore";

const net = createMemoryNet();

/** What iroh hands back. A ticket names the host, so the host holds a real one. */
const ALEX = "a".repeat(64);
const SAM = "b".repeat(64);
const STRANGER = "c".repeat(64);

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

/** Time the test owns, so a backoff or a deadline is a step rather than a wait. */
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
    };
}

beforeEach(() => {
    net.reset();
    forgetRoundPeers();
    useFlowStore.setState({ collabEnabled: true, collabRelayEnabled: true });
    shared = makeFlowRound({});
});

describe("startCollabSession", () => {
    it("listens on the local endpoint", async () => {
        const session = await open(ALEX);
        expect(session!.endpointId).toBe(ALEX);
        expect(session!.roundId).toBe(shared.id);
        expect(net.calls.map((c) => c.op)).toContain("listen");
    });

    it("keeps running when a known peer cannot be reached", async () => {
        const session = await open(ALEX, { dial: ["gone"] });
        expect(session).not.toBeNull();
        expect(session!.peers()).toEqual([]);
    });

    it("re-dials a known peer with no ticket, which is what resume does", async () => {
        // The host already knows sam, the way a sidecar's peer list says it does.
        const host = await open(ALEX, { dial: ["sam"] });
        const guest = await open("sam", { dial: [ALEX] });
        await settle();
        expect(guest!.peers().map((p) => p.endpointId)).toEqual([ALEX]);
        expect(host!.peers().map((p) => p.endpointId)).toEqual(["sam"]);
    });

    it("reports the peer list as it changes", async () => {
        const seen: CollabPeer[][] = [];
        const host = await open(ALEX, {
            dial: ["sam"],
            onPeersChanged: (peers: CollabPeer[]) => seen.push(peers),
        });
        await open("sam", { dial: [ALEX] });
        await settle();
        expect(seen.at(-1)!.map((p) => p.endpointId)).toEqual(["sam"]);
        expect(host!.peers()).toHaveLength(1);
    });

    it("drops a peer from both lists when the link closes", async () => {
        const host = await open(ALEX, { dial: ["sam"] });
        const guest = await open("sam", { dial: [ALEX] });
        await settle();
        await guest!.stop();
        await settle();
        expect(host!.peers()).toEqual([]);
        expect(guest!.peers()).toEqual([]);
    });

    it("stops the link it started", async () => {
        const session = await open(ALEX);
        await session!.stop();
        expect(net.calls.map((c) => c.op)).toContain("stop");
    });

    it("mints a ticket that names this host and this round", async () => {
        const session = await open(ALEX);
        const ticket = session!.share("partner");
        expect(ticket).toMatchObject({
            endpointId: ALEX,
            roundId: shared.id,
            role: "partner",
            relay: true,
        });
        expect(encodeTicket(ticket)).toContain("ebb1:");
    });

    it("mints a fresh ticket each time, replacing the unspent one", async () => {
        const session = await open(ALEX);
        expect(session!.share("partner").secret).not.toBe(session!.share("partner").secret);
    });

    it("carries the relay stance the settings hold into the ticket", async () => {
        useFlowStore.setState({ collabRelayEnabled: false });
        const session = await open(ALEX);
        expect(session!.share("partner").relay).toBe(false);
    });
});

describe("a link that drops mid-round", () => {
    /** A link that hands back every connection it dials, so a test can cut one. */
    function watched(endpointId: string, dialled: PeerConn[]): PeerLinkFactory {
        return async (config) => {
            const link = await net.create(endpointId)(config);
            return {
                ...link,
                async dial(target: string) {
                    const conn = await link.dial(target);
                    dialled.push(conn);
                    return conn;
                },
            };
        };
    }

    // The dial that opened a link only retried while the session was coming
    // up, and only when a test handed it a scheduler. A wifi blip mid-round
    // left the peer gone for the rest of the round, and the debater's only way
    // back was to close the flow and open it again.
    it("dials the peer again, without anyone being asked", async () => {
        const clock = manualClock();
        const conns: PeerConn[] = [];
        const host = (await open(ALEX))!;
        const guest = (await open("sam", {
            createLink: watched("sam", conns),
            ticket: encodeTicket(host.share("partner")),
            dial: [ALEX],
            schedule: clock.schedule,
        }))!;
        await settle();
        expect(guest.peers()).toHaveLength(1);

        conns[0].close();
        await settle();
        expect(guest.peers()).toHaveLength(0);

        // The backoff comes round and the guest reaches the host again, with
        // no ticket and nothing on screen to answer.
        for (let i = 0; i < 6 && guest.peers().length === 0; i++) {
            clock.advance(60_000);
            await settle();
        }
        expect(guest.peers()).toHaveLength(1);
        expect(conns.length).toBeGreaterThan(1);

        await host.stop();
        await guest.stop();
    });

    it("stops trying once the session is over", async () => {
        const clock = manualClock();
        const conns: PeerConn[] = [];
        const host = (await open(ALEX))!;
        const guest = (await open("sam", {
            createLink: watched("sam", conns),
            ticket: encodeTicket(host.share("partner")),
            dial: [ALEX],
            schedule: clock.schedule,
        }))!;
        await settle();

        await guest.stop();
        const dialsAtStop = conns.length;
        for (let i = 0; i < 4; i++) {
            clock.advance(60_000);
            await settle();
        }
        expect(conns).toHaveLength(dialsAtStop);
        expect(guest.peers()).toHaveLength(0);

        await host.stop();
    });

    it("says it is reconnecting from the drop until the peer answers again", async () => {
        const clock = manualClock();
        const conns: PeerConn[] = [];
        const host = (await open(ALEX))!;
        const guest = (await open("sam", {
            createLink: watched("sam", conns),
            ticket: encodeTicket(host.share("partner")),
            dial: [ALEX],
            schedule: clock.schedule,
        }))!;
        await settle();
        expect(guest.reconnecting()).toBe(false);

        conns[0].close();
        await settle();
        expect(guest.reconnecting()).toBe(true);

        for (let i = 0; i < 6 && guest.peers().length === 0; i++) {
            clock.advance(60_000);
            await settle();
        }
        expect(guest.peers()).toHaveLength(1);
        expect(guest.reconnecting()).toBe(false);

        await host.stop();
        await guest.stop();
    });

    // A backoff is up to half a minute wide, so a session that ended while one
    // was armed would dial a peer long after the round was closed.
    it("cancels a backoff that outlived the session", async () => {
        const clock = manualClock();
        const conns: PeerConn[] = [];
        const host = (await open(ALEX))!;
        const guest = (await open("sam", {
            createLink: watched("sam", conns),
            ticket: encodeTicket(host.share("partner")),
            dial: [ALEX],
            schedule: clock.schedule,
        }))!;
        await settle();
        conns[0].close();
        await settle();
        expect(guest.reconnecting()).toBe(true);

        await guest.stop();
        expect(guest.reconnecting()).toBe(false);
        const dialsAtStop = conns.length;
        clock.advance(60_000);
        await settle();
        expect(conns).toHaveLength(dialsAtStop);

        await host.stop();
    });

    it("keeps a disconnected peer gone, however long the backoff runs", async () => {
        const clock = manualClock();
        const conns: PeerConn[] = [];
        const host = (await open(ALEX))!;
        const guest = (await open("sam", {
            createLink: watched("sam", conns),
            ticket: encodeTicket(host.share("partner")),
            dial: [ALEX],
            schedule: clock.schedule,
        }))!;
        await settle();
        expect(guest.peers()).toHaveLength(1);

        guest.disconnect(ALEX);
        await settle();
        expect(guest.peers()).toEqual([]);
        // The link really went, so the host is not left holding a peer that
        // walked away.
        expect(host.peers()).toEqual([]);
        expect(guest.reconnecting()).toBe(false);

        const dialsAtDisconnect = conns.length;
        for (let i = 0; i < 6; i++) {
            clock.advance(60_000);
            await settle();
        }
        expect(conns).toHaveLength(dialsAtDisconnect);
        expect(guest.peers()).toEqual([]);

        await host.stop();
        await guest.stop();
    });

    // The link dropped, the ladder is climbing, and the debater decides they
    // are done with that peer. Nothing about that is worth another dial.
    it("stops a backoff already climbing when the debater disconnects", async () => {
        const clock = manualClock();
        const conns: PeerConn[] = [];
        const host = (await open(ALEX))!;
        const guest = (await open("sam", {
            createLink: watched("sam", conns),
            ticket: encodeTicket(host.share("partner")),
            dial: [ALEX],
            schedule: clock.schedule,
        }))!;
        await settle();
        conns[0].close();
        await settle();
        expect(guest.reconnecting()).toBe(true);

        guest.disconnect(ALEX);
        expect(guest.reconnecting()).toBe(false);

        const dialsAtDisconnect = conns.length;
        for (let i = 0; i < 6; i++) {
            clock.advance(60_000);
            await settle();
        }
        expect(conns).toHaveLength(dialsAtDisconnect);
        expect(guest.peers()).toEqual([]);

        await host.stop();
        await guest.stop();
    });

    it("does not let a disconnected peer dial its way back in", async () => {
        const clock = manualClock();
        const host = (await open(ALEX))!;
        const guest = (await open("sam", {
            ticket: encodeTicket(host.share("partner")),
            dial: [ALEX],
            schedule: clock.schedule,
        }))!;
        await settle();
        expect(host.peers()).toHaveLength(1);

        host.disconnect("sam");
        await settle();
        expect(host.peers()).toEqual([]);

        // A known peer needs no ticket, which is exactly what makes the
        // disconnect worth enforcing on the way in.
        const again = (await open("sam", { dial: [ALEX], schedule: clock.schedule }))!;
        await settle();
        expect(host.peers()).toEqual([]);
        expect(again.peers()).toEqual([]);

        await host.stop();
        await guest.stop();
        await again.stop();
    });

    // The cut used to live in the session closure alone, while the round's own
    // peer list - which reaches the sidecar and comes back off it - only ever
    // grew. The next open dialled the peer again and admitted them on
    // membership, so Disconnect read as permanent and was a per-session mute.
    it("keeps a disconnected peer out of what the round remembers, so the cut survives the next open", async () => {
        const clock = manualClock();
        // What opening a round does before a session starts: the record exists
        // and is empty, because no sidecar for it does.
        setRoundPeers(shared.id, [], []);
        const host = (await open(ALEX, { schedule: clock.schedule }))!;
        const guest = (await open("sam", {
            ticket: encodeTicket(host.share("partner")),
            dial: [ALEX],
            schedule: clock.schedule,
        }))!;
        await settle();
        expect(host.peers()).toHaveLength(1);
        expect(knownRoundPeers(shared.id)).toEqual(["sam"]);

        host.disconnect("sam");
        await settle();
        expect(knownRoundPeers(shared.id)).toEqual([]);
        await host.stop();
        await guest.stop();

        // What the next open is: a fresh session dialling and admitting off the
        // round's record. Nothing there names the peer, and they hold no
        // ticket, so the refusal is silent on both paths.
        const reopened = (await open(ALEX, {
            dial: knownRoundPeers(shared.id),
            schedule: clock.schedule,
        }))!;
        const back = (await open("sam", { dial: [ALEX], schedule: clock.schedule }))!;
        await settle();
        expect(net.calls.filter((c) => c.op === "dial" && c.endpointId === "sam")).toEqual([]);
        expect(reopened.peers()).toEqual([]);
        expect(back.peers()).toEqual([]);

        await reopened.stop();
        await back.stop();
    });

    // Resume is symmetric, so both sides reach out. Two connections landing in
    // one slot left whichever lost the map entry open, unreachable, unclosed,
    // and still counted as a peer by the far side.
    it("keeps one connection when both sides reach each other at once", async () => {
        const hostConns: PeerConn[] = [];
        const guestConns: PeerConn[] = [];
        const host = (await open(ALEX, { createLink: watched(ALEX, hostConns) }))!;
        const guest = (await open("sam", {
            createLink: watched("sam", guestConns),
            ticket: encodeTicket(host.share("partner")),
            dial: [ALEX],
        }))!;
        await settle();
        expect(host.peers()).toHaveLength(1);
        expect(guestConns).toHaveLength(1);

        let cut = 0;
        guestConns[0].onClose(() => cut++);

        // The host reaches for a guest it is already holding, which is what a
        // contact invited onto a round they have just joined looks like.
        await host.invite("sam");
        await settle();
        expect(hostConns).toHaveLength(1);

        expect(host.peers().map((p) => p.endpointId)).toEqual(["sam"]);
        expect(guest.peers().map((p) => p.endpointId)).toEqual([ALEX]);
        // Both ends dropped the same one, so the guest's own dial is closed
        // and the peer is still there on the connection that survived.
        expect(cut).toBe(1);

        // And it is the same connection on both sides: the guest leaving is
        // something the host hears about.
        await guest.stop();
        await settle();
        expect(host.peers()).toEqual([]);
        await host.stop();
    });
});

describe("what a dialler is told", () => {
    function hello(over: Partial<Extract<WireMessage, { type: "hello" }>> = {}): WireMessage {
        return {
            type: "hello",
            protocol: PROTOCOL_MAJOR,
            app: "0.11.0",
            endpointId: STRANGER,
            roundId: shared.id,
            role: "partner",
            capabilities: [],
            ...over,
        };
    }

    /** Dials the host by hand, so a test sees exactly what comes back. */
    async function knock(from: string, msg: WireMessage) {
        const link = await net.create(from)({ discovery: "mdns", relay: true });
        const conn = await link.dial(ALEX);
        const answers: WireMessage[] = [];
        let closed = false;
        conn.onMessage((m) => answers.push(m));
        conn.onClose(() => {
            closed = true;
        });
        conn.send(msg);
        await settle();
        return { conn, answers, closed };
    }

    // An EndpointId is permanent and every peer who ever shared with this
    // install holds one, so a stranger who dials learns that something closed
    // and nothing else at all.
    it("puts nothing on the wire for a refusal it is not meant to see", async () => {
        await open(ALEX);
        const { answers, closed } = await knock(STRANGER, hello());
        expect(answers).toEqual([]);
        expect(closed).toBe(true);
    });

    it("tells a stranger on another version nothing about this one", async () => {
        await open(ALEX);
        const { answers } = await knock(STRANGER, hello({ protocol: PROTOCOL_MAJOR + 1 }));
        expect(answers).toEqual([]);
    });

    it("names a skew to a caller holding the ticket, without naming a version", async () => {
        const host = (await open(ALEX))!;
        const ticket = host.share("partner");
        const { answers } = await knock(
            SAM,
            hello({ endpointId: SAM, protocol: PROTOCOL_MAJOR + 1, ticket: ticket.secret }),
        );
        expect(answers).toEqual([{ type: "helloAck", ok: false, reason: VERSION_SKEW }]);
        expect(JSON.stringify(answers)).not.toContain("0.11.0");
    });

    // The refusing side wrote that string, and it lands on this side's screen.
    it("never repeats the words a refusing host chose", async () => {
        const link = await net.create(ALEX)({ discovery: "mdns", relay: true });
        await link.listen((conn) => {
            conn.onMessage(() => {
                conn.send({
                    type: "helloAck",
                    ok: false,
                    reason: "ebb says: your flow is corrupt, call 555-0100 to recover it",
                });
                conn.close();
            });
        });

        const guest = (await open(SAM))!;
        const err = await guest.invite(ALEX).then(
            () => null,
            (e: unknown) => e as Error,
        );
        expect(err!.message).toBe("That peer refused the connection");
        expect(err!.message).not.toContain("555");
    });
});

describe("a dialler that never greets", () => {
    /** Opens a connection to the host and says nothing at all on it. */
    async function silent(from: string) {
        const link = await net.create(from)({ discovery: "mdns", relay: true });
        const conn = await link.dial(ALEX);
        const state = { closed: false };
        conn.onClose(() => {
            state.closed = true;
        });
        return state;
    }

    // Every refusal in the admission path is inside the greeting handler, so a
    // connection that never enters it was never refused and never closed. A
    // stranger who knows the EndpointId could hold one slot per dial, for the
    // whole round, with nothing on the debater's screen to say so.
    it("is closed once the deadline passes", async () => {
        const clock = manualClock();
        const host = (await open(ALEX, { schedule: clock.schedule }))!;
        const stranger = await silent(STRANGER);
        await settle();
        expect(stranger.closed).toBe(false);

        clock.advance(HANDSHAKE_MS);
        await settle();
        expect(stranger.closed).toBe(true);
        // Nothing was admitted, so nothing was on the peer list to lose.
        expect(host.peers()).toEqual([]);

        await host.stop();
    });

    it("does not take an admitted peer with it when the deadline comes round", async () => {
        const clock = manualClock();
        const host = (await open(ALEX, { schedule: clock.schedule }))!;
        const guest = (await open(SAM, {
            ticket: encodeTicket(host.share("partner")),
            dial: [ALEX],
            schedule: clock.schedule,
        }))!;
        await settle();
        expect(host.peers()).toHaveLength(1);

        clock.advance(HANDSHAKE_MS * 2);
        await settle();
        expect(host.peers()).toHaveLength(1);
        expect(guest.peers()).toHaveLength(1);

        await host.stop();
        await guest.stop();
    });
});

describe("a peer's claim on a cell", () => {
    /** A guest admitted by ticket, whose connection the test speaks over. */
    async function admitted(host: CollabSession): Promise<PeerConn> {
        const secret = host.share("partner").secret;
        const link = await net.create(SAM)({ discovery: "mdns", relay: true });
        const conn = await link.dial(ALEX);
        conn.send({
            type: "hello",
            protocol: PROTOCOL_MAJOR,
            app: "0.11.0",
            endpointId: SAM,
            roundId: shared.id,
            role: "partner",
            capabilities: [],
            ticket: secret,
        });
        await settle();
        return conn;
    }

    // The cell goes straight into the presence table, where a row nobody can
    // hold would sit unmatched for the rest of the round.
    it("ignores a cell that is not one", async () => {
        const host = (await open(ALEX))!;
        const conn = await admitted(host);
        expect(host.peers()).toHaveLength(1);
        expect(getPresences()).toEqual([]);

        for (const cell of [
            { sheetId: "sheet_1", col: -1, row: 0 },
            { sheetId: "sheet_1", col: 0, row: 1.5 },
            { sheetId: 7, col: 0, row: 0 },
            { sheetId: "sheet_1", col: "0", row: 0 },
            { sheetId: "sheet_1", row: 0 },
            "sheet_1",
        ]) {
            conn.send({ type: "presence", cell } as WireMessage);
            conn.send({ type: "cursor", cell } as WireMessage);
            await settle();
            expect(getPresences()).toEqual([]);
        }

        conn.send({ type: "presence", cell: { sheetId: "sheet_1", col: 1, row: 2 } });
        await settle();
        expect(getPresences()).toEqual([
            {
                endpointId: SAM,
                sheetId: "sheet_1",
                col: 1,
                row: 2,
                heldAt: expect.any(Number),
                editing: true,
            },
        ]);
        await host.stop();
    });

    // A cursor is not a claim. Painting it is the point; refusing a keystroke
    // on it would make a partner reading over your shoulder cost you a cell.
    it("records a resting cursor without claiming the cell", async () => {
        const host = (await open(ALEX))!;
        const conn = await admitted(host);

        conn.send({ type: "cursor", cell: { sheetId: "sheet_1", col: 1, row: 2 } });
        await settle();
        expect(getPresences()).toEqual([
            {
                endpointId: SAM,
                sheetId: "sheet_1",
                col: 1,
                row: 2,
                heldAt: expect.any(Number),
                editing: false,
            },
        ]);

        // One entry per peer either way round: a cursor that started editing
        // is the same peer in the same place, not a second mark.
        conn.send({ type: "presence", cell: { sheetId: "sheet_1", col: 1, row: 2 } });
        await settle();
        expect(getPresences()).toHaveLength(1);
        expect(getPresences()[0].editing).toBe(true);

        conn.send({ type: "cursor", cell: null });
        await settle();
        expect(getPresences()).toEqual([]);
        await host.stop();
    });
});

/**
 * Every window hears every accepted connection, because the round it belongs
 * to arrives in the hello and the shell reads no further than the bytes. So
 * an accepted connection is nobody's until a window says the peer is theirs,
 * and admitting them is the one thing that says it. Writing is not: a window
 * with a different flow open answers the same hello with a refusal, and
 * latching on that write left the refusing window holding a guest it then hung
 * up on, with the host's own ack refused as another window's.
 */
describe("what admitting a peer tells the shell", () => {
    /** The memory transport with the shell's claim on it, which only the
     *  desktop adapter has. */
    function claiming(endpointId: string, claimed: string[]): PeerLinkFactory {
        return async (config) => {
            const link = await net.create(endpointId)(config);
            return {
                ...link,
                async listen(onPeer: (conn: PeerConn) => void) {
                    await link.listen((conn) =>
                        onPeer({ ...conn, claim: () => claimed.push(conn.id) }),
                    );
                },
            };
        };
    }

    /** Dials the host with a hello and lets the handshake run out. */
    async function greet(from: string, msg: WireMessage): Promise<void> {
        const link = await net.create(from)({ discovery: "mdns", relay: true });
        const conn = await link.dial(ALEX);
        conn.send(msg);
        await settle();
    }

    it("claims the connection of a peer it lets in", async () => {
        const claimed: string[] = [];
        const host = (await open(ALEX, { createLink: claiming(ALEX, claimed) }))!;
        const secret = host.share("partner").secret;

        await greet(SAM, {
            type: "hello",
            protocol: PROTOCOL_MAJOR,
            app: "0.11.0",
            endpointId: SAM,
            roundId: shared.id,
            role: "partner",
            capabilities: [],
            ticket: secret,
        });

        expect(host.peers()).toHaveLength(1);
        expect(claimed).toEqual([SAM]);
    });

    it("claims nothing from a peer it refuses", async () => {
        const claimed: string[] = [];
        const host = (await open(ALEX, { createLink: claiming(ALEX, claimed) }))!;

        await greet(STRANGER, {
            type: "hello",
            protocol: PROTOCOL_MAJOR,
            app: "0.11.0",
            endpointId: STRANGER,
            roundId: shared.id,
            role: "partner",
            capabilities: [],
        });

        expect(host.peers()).toEqual([]);
        expect(claimed).toEqual([]);
    });
});
