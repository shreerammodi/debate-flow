import { beforeEach, describe, expect, it } from "vitest";

import type { Contacts } from "@/lib/collab/contacts";
import { helloFrom } from "@/lib/collab/handshake";
import { INVITED, type InviteNotice } from "@/lib/collab/invite";
import { startInviteListener } from "@/lib/collab/inviteListener";
import type { PeerConn, PeerLinkFactory, WireMessage } from "@/lib/collab/peerLink";
import { createMemoryNet } from "@/lib/collab/peerLinkMemory";
import { useFlowStore } from "@/lib/store/useFlowStore";

const ALEX = "alex";
const STRANGER = "who";

const net = createMemoryNet();
const contacts: Contacts = { [ALEX]: { name: "Alex", role: "partner" } };

let heard: InviteNotice[];
let openRound: string | null;
/** The pending handshake deadline, so a test can reach it without a clock. */
let deadline: (() => void) | null;

function listener(table: Contacts = contacts, createLink: PeerLinkFactory = net.create("me")) {
    return startInviteListener({
        createLink,
        contacts: () => table,
        openRoundId: () => openRound,
        onInvite: (notice) => heard.push(notice),
        schedule: (fn) => {
            deadline = fn;
            return () => {
                deadline = null;
            };
        },
    });
}

interface Dial {
    answers: WireMessage[];
    closed(): boolean;
    send(msg: WireMessage): void;
}

/** Opens a connection to the listener and says nothing on it. */
async function dial(from: string): Promise<Dial> {
    const link = await net.create(from)({ discovery: "mdns", relay: true });
    const conn = await link.dial("me");
    const answers: WireMessage[] = [];
    let closed = false;
    conn.onMessage((msg) => answers.push(msg));
    conn.onClose(() => {
        closed = true;
    });
    return { answers, closed: () => closed, send: (msg) => conn.send(msg) };
}

/** Dials the listener the way a partner offering a round does. */
async function offer(from: string, label: string, roundId = "their-round"): Promise<Dial> {
    const d = await dial(from);
    d.send(
        helloFrom({
            endpointId: from,
            roundId,
            role: "partner",
            appVersion: "0.11.0",
            label,
        }),
    );
    await Promise.resolve();
    return d;
}

beforeEach(() => {
    net.reset();
    heard = [];
    openRound = null;
    deadline = null;
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
        const { answers } = await offer(ALEX, "Round 3");
        expect(answers).toEqual([{ type: "helloAck", ok: false, reason: INVITED }]);
    });

    it("says nothing at all to a peer nobody saved", async () => {
        await listener();
        const { answers } = await offer(STRANGER, "Round 3");
        expect(heard).toEqual([]);
        expect(answers).toEqual([]);
    });

    it("joins nothing on its own", async () => {
        // The round only lands when the debater says so, so the listener
        // never asks for state.
        await listener();
        const { answers } = await offer(ALEX, "Round 3");
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

    // Answering is what makes the connection this window's: a notice went out,
    // so releasing it cannot be releasing a peer another window admitted.
    it("hangs up on the contact it answered", async () => {
        await listener();
        const alex = await offer(ALEX, "Round 3");
        expect(alex.closed()).toBe(true);
    });

    // The shell hands one accepted connection to every window, so an immediate
    // hang-up on a hello this window did not answer could be a hang-up on a
    // peer another window was in the middle of admitting.
    it("leaves a peer it never answered alone until the deadline", async () => {
        await listener();
        const stranger = await offer(STRANGER, "Round 3");
        expect(stranger.closed()).toBe(false);

        deadline!();

        expect(stranger.closed()).toBe(true);
        expect(stranger.answers).toEqual([]);
    });

    // A stranger who dials and says nothing at all would otherwise hold the
    // connection, and its tasks, for as long as ebb runs.
    it("releases a dial that says nothing at all", async () => {
        await listener();
        const quiet = await dial(STRANGER);
        expect(quiet.closed()).toBe(false);

        deadline!();

        expect(quiet.closed()).toBe(true);
    });

    // Writing again must not buy a longer stay. The shell addresses an owned
    // connection's messages to the window that claimed it, so a second line
    // reaching a window holding no session is a peer nobody admitted - the
    // case the deadline is for. A stranger who greets and then chatters would
    // otherwise pin one inbound slot per dial for the life of the process.
    it("releases a stranger that keeps talking after its greeting", async () => {
        await listener();
        const stranger = await offer(STRANGER, "Round 3");
        stranger.send({ type: "vector", seen: {} });
        await Promise.resolve();

        deadline!();

        expect(stranger.closed()).toBe(true);
        expect(stranger.answers).toEqual([]);
    });

    // The same for a saved contact arriving about the round this window has
    // open: this window answers nothing, so it is holding a connection it is
    // not using. Closing one the shell says belongs to another window is
    // refused there and costs this window only its own handle.
    it("releases a contact it left for another window, however much they say", async () => {
        openRound = "my-round";
        await listener();
        const peer = await offer(ALEX, "Round 3", "my-round");
        peer.send({ type: "vector", seen: {} });
        await Promise.resolve();

        deadline!();

        expect(peer.closed()).toBe(true);
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

/**
 * One accepted connection reaches every window, so a hello about the round
 * this window has open is somebody else's peer joining it - the window that is
 * sharing that round answers, and this one must not put a notice on screen or
 * refuse on its behalf.
 */
describe("a peer arriving about the round this window has open", () => {
    beforeEach(() => {
        openRound = "my-round";
    });

    it("draws no notice and no answer, even from a saved contact", async () => {
        await listener();
        const { answers } = await offer(ALEX, "Round 3", "my-round");
        expect(heard).toEqual([]);
        expect(answers).toEqual([]);
    });

    it("still hears an offer of any other round", async () => {
        await listener();
        const { answers } = await offer(ALEX, "Round 3", "their-round");
        expect(heard.map((n) => n.roundId)).toEqual(["their-round"]);
        expect(answers).toEqual([{ type: "helloAck", ok: false, reason: INVITED }]);
    });
});

/**
 * The shell keeps an accepted connection ownerless until a window says the
 * peer is theirs, because the round arrives in the hello and the shell reads
 * no further than the bytes. This listener answers connections it is refusing,
 * and a refusal that took ownership would hand the guest the hosting window
 * just admitted to the window that hung up on them.
 */
describe("what refusing an invite tells the shell", () => {
    /** The memory transport with the shell's claim on it, which only the
     *  desktop adapter has. */
    function claiming(claimed: string[]): PeerLinkFactory {
        return async (config) => {
            const link = await net.create("me")(config);
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

    it("answers and hangs up without taking the connection", async () => {
        const claimed: string[] = [];
        await listener(contacts, claiming(claimed));
        const { answers, closed } = await offer(ALEX, "Round 3");

        expect(answers).toEqual([{ type: "helloAck", ok: false, reason: INVITED }]);
        expect(closed()).toBe(true);
        expect(claimed).toEqual([]);
    });

    it("takes nothing from a stranger it never answers either", async () => {
        const claimed: string[] = [];
        await listener(contacts, claiming(claimed));
        await offer(STRANGER, "Round 3");

        expect(claimed).toEqual([]);
    });
});
