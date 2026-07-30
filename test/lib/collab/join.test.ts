import { beforeEach, describe, expect, it } from "vitest";

import { seedDoc } from "@/lib/collab/doc";
import { joinRound } from "@/lib/collab/join";
import { merge } from "@/lib/collab/merge";
import { applyOp, type OpContext } from "@/lib/collab/ops";
import type { PeerLinkFactory, WireMessage } from "@/lib/collab/peerLink";
import { createMemoryNet } from "@/lib/collab/peerLinkMemory";
import { recoverReplica } from "@/lib/collab/persist";
import { getReplica } from "@/lib/collab/replica";
import { forgetRoundPeers, knownRoundPeers, setRoundPeers } from "@/lib/collab/roundPeers";
import { applyRemoteDoc } from "@/lib/collab/runtime";
import { startCollabSession } from "@/lib/collab/session";
import { setSidecarFs } from "@/lib/collab/sidecarFs";
import { createSidecarFs } from "@/lib/collab/sidecarFsMemory";
import { createClock } from "@/lib/collab/stamp";
import { encodeTicket } from "@/lib/collab/ticket";
import type { CollabDoc } from "@/lib/collab/types";
import { makeFlowRound, type FlowRound } from "@/lib/model/flow";
import { serializeFlow } from "@/lib/persistence/flowFile";
import { parseFlowFile } from "@/lib/persistence/flowFile";
import type { FlowFs } from "@/lib/persistence/flowFs";
import { createFlowFs } from "@/lib/persistence/flowFsMemory";
import { saveRecents } from "@/lib/persistence/recents";
import { useFlowStore } from "@/lib/store/useFlowStore";

const net = createMemoryNet();

/** What iroh hands back. A ticket names the host, so the host holds a real one. */
const ALEX = "a".repeat(64);

let shared: FlowRound;
let fs: FlowFs;

function side(base: FlowRound) {
    let doc = seedDoc(base);
    let t = 1_000;
    const ctx: OpContext = { actor: ALEX, clock: createClock(ALEX, () => t++) };
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

async function hostWithTicket(): Promise<string> {
    const host = (await startCollabSession({
        createLink: net.create(ALEX),
        roundId: shared.id,
        appVersion: "0.11.0",
        ...side(shared),
    }))!;
    return encodeTicket(host.share("partner"));
}

beforeEach(async () => {
    // A join is a route that starts a session, so it asks collabLive(), which
    // is the desktop shell as well as the switch. The suite drives the protocol
    // over an in-process transport and has to stand in for the shell - and then
    // pin the sidecar to memory, because the port picks its adapter off the
    // same signal and the Tauri one has no shell to invoke here.
    (window as unknown as Record<string, unknown>).__TAURI_INTERNALS__ = {};
    setSidecarFs(createSidecarFs());
    net.reset();
    forgetRoundPeers();
    useFlowStore.setState({
        collabEnabled: true,
        collabRelayEnabled: true,
        collabName: "Rin",
    });
    shared = makeFlowRound({});
    shared.sheets.find((s) => s.kind !== "cx")!.data = [["perm", "link"]];
    fs = createFlowFs();
});

describe("joinRound", () => {
    it("writes a real file of its own for a round it has never seen", async () => {
        const ticket = await hostWithTicket();
        const result = await joinRound({
            ticket,
            createLink: net.create("sam"),
            appVersion: "0.11.0",
            fs,
        });
        expect(result!.created).toBe(true);
        expect(result!.roundId).toBe(shared.id);
        expect(result!.hostEndpointId).toBe(ALEX);

        const written = await fs.readFlow(result!.path);
        expect(written).not.toBeNull();
        expect(written!.text).toContain("perm");
    });

    // The host hears this before anything else, and offers it as the contact
    // name. A nameless greeting is saved on the far side as a short
    // EndpointId, and the round's own session re-dialling later is too late.
    it("tells the host what to call this side", async () => {
        const greetings: WireMessage[] = [];
        const host = (await startCollabSession({
            createLink: net.create(ALEX),
            roundId: shared.id,
            appVersion: "0.11.0",
            ...side(shared),
            contacts: () => ({}),
        }))!;
        const ticket = encodeTicket(host.share("partner"));

        const watched: PeerLinkFactory = async (config) => {
            const link = await net.create("sam")(config);
            return {
                ...link,
                async dial(target: string) {
                    const conn = await link.dial(target);
                    const send = conn.send.bind(conn);
                    return {
                        ...conn,
                        send(msg: WireMessage) {
                            greetings.push(msg);
                            send(msg);
                        },
                    };
                },
            };
        };

        await joinRound({ ticket, createLink: watched, appVersion: "0.11.0", fs });

        const hello = greetings.find((m) => m.type === "hello");
        expect(hello).toMatchObject({ type: "hello", name: "Rin" });
    });

    it("syncs into the file it already has rather than making a duplicate", async () => {
        const ticket = await hostWithTicket();
        // The guest already owns this round from an earlier session.
        const dir = (await fs.locations()).flowsDir;
        const existing = await fs.createFlow(dir, "round.ebb", serializeFlow(shared));
        await saveRecents(fs, [{ path: existing, openedAt: 1 }]);

        const result = await joinRound({
            ticket,
            createLink: net.create("sam"),
            appVersion: "0.11.0",
            fs,
        });
        expect(result!.created).toBe(false);
        expect(result!.path).toBe(existing);
    });

    it("refuses a string that is not a ticket, before it touches the network", async () => {
        await expect(
            joinRound({
                ticket: "hello",
                createLink: net.create("sam"),
                appVersion: "0.11.0",
                fs,
            }),
        ).rejects.toThrow(/ticket/i);
        expect(net.calls).toEqual([]);
    });

    it("surfaces the reason when the host refuses", async () => {
        const ticket = await hostWithTicket();
        // Spend the ticket, so the second presentation is refused.
        await joinRound({ ticket, createLink: net.create("sam"), appVersion: "0.11.0", fs });
        await expect(
            joinRound({
                ticket,
                createLink: net.create("kim"),
                appVersion: "0.11.0",
                fs,
            }),
        ).rejects.toThrow();
    });

    // A ticket names one round. The host scopes every hello to the round it is
    // holding; the guest had no counterpart, so it took whatever came back and
    // keyed the file, the peer record and the navigation off that.
    it("refuses a document for a round the ticket did not name", async () => {
        const link = await net.create(ALEX)({ discovery: "mdns", relay: true });
        const elsewhere = seedDoc(makeFlowRound({}));
        await link.listen((conn) => {
            conn.onMessage(() => {
                conn.send({ type: "state", doc: elsewhere });
                conn.close();
            });
        });
        // A round is open in this window, and its peer set is module-global.
        setRoundPeers("round_open", ["kim"]);

        await expect(
            joinRound({
                ticket: encodeTicket({
                    endpointId: ALEX,
                    roundId: shared.id,
                    role: "partner",
                    secret: "s".repeat(24),
                    relay: true,
                }),
                createLink: net.create("sam"),
                appVersion: "0.11.0",
                fs,
            }),
        ).rejects.toThrow(/hung up/i);
        // Neither the round it offered nor the round that is open picked up a
        // peer, and the open one still knows its own.
        expect(knownRoundPeers(elsewhere.roundId)).toEqual([]);
        expect(knownRoundPeers("round_open")).toEqual(["kim"]);
    });

    it("reaches no network at all while shared editing is off", async () => {
        useFlowStore.setState({ collabEnabled: false });
        const result = await joinRound({
            ticket: encodeTicket({
                endpointId: ALEX,
                roundId: "r",
                role: "partner",
                secret: "s".repeat(24),
                relay: false,
            }),
            createLink: net.create("sam"),
            appVersion: "0.11.0",
            fs,
        });
        expect(result).toBeNull();
        expect(net.calls).toEqual([]);
    });

    it("closes the link it opened, so the round's own session owns the peer", async () => {
        const ticket = await hostWithTicket();
        await joinRound({ ticket, createLink: net.create("sam"), appVersion: "0.11.0", fs });
        expect(net.calls.filter((c) => c.op === "stop").length).toBeGreaterThan(0);
    });
});

describe("joining a contact's invitation", () => {
    /** Alex, having invited Sam: the dial put Sam in the known list. */
    async function hostWhoInvited(): Promise<void> {
        const host = (await startCollabSession({
            createLink: net.create(ALEX),
            roundId: shared.id,
            appVersion: "0.11.0",
            ...side(shared),
        }))!;
        await host.invite("sam").catch(() => {
            // Sam has nothing bound to answer with, which is the state a
            // debater is in between the corner message and the Join.
        });
    }

    it("takes the round with no ticket at all", async () => {
        await hostWhoInvited();
        const result = await joinRound({
            invite: { endpointId: ALEX, roundId: shared.id },
            createLink: net.create("sam"),
            appVersion: "0.11.0",
            fs,
        });
        expect(result!.created).toBe(true);
        expect(result!.hostEndpointId).toBe(ALEX);
        expect((await fs.readFlow(result!.path))!.text).toContain("perm");
    });

    it("is refused by a host who never invited this peer", async () => {
        await startCollabSession({
            createLink: net.create(ALEX),
            roundId: shared.id,
            appVersion: "0.11.0",
            ...side(shared),
        });
        await expect(
            joinRound({
                invite: { endpointId: ALEX, roundId: shared.id },
                createLink: net.create("sam"),
                appVersion: "0.11.0",
                fs,
            }),
        ).rejects.toThrow();
    });

    it("reaches no network at all while shared editing is off", async () => {
        useFlowStore.setState({ collabEnabled: false });
        const result = await joinRound({
            invite: { endpointId: ALEX, roundId: shared.id },
            createLink: net.create("sam"),
            appVersion: "0.11.0",
            fs,
        });
        expect(result).toBeNull();
        expect(net.calls).toEqual([]);
    });

    it("refuses a call with neither a ticket nor an invitation", async () => {
        await expect(
            joinRound({ createLink: net.create("sam"), appVersion: "0.11.0", fs }),
        ).rejects.toThrow(/ticket/i);
        expect(net.calls).toEqual([]);
    });
});

describe("what a joined round remembers", () => {
    it("keeps the host, so opening the file re-dials them with no ticket", async () => {
        const ticket = await hostWithTicket();
        const result = await joinRound({
            ticket,
            createLink: net.create("sam"),
            appVersion: "0.11.0",
            fs,
        });
        expect(knownRoundPeers(result!.roundId)).toEqual([ALEX]);
    });

    it("keeps the host's peers across the open", async () => {
        const ticket = await hostWithTicket();
        const joined = await joinRound({
            ticket,
            createLink: net.create("sam"),
            appVersion: "0.11.0",
            fs,
        });
        const round = parseFlowFile((await fs.readFlow(joined!.path))!.text);
        expect(await recoverReplica(round, serializeFlow(round))).toEqual([ALEX]);
    });

    it("adopts the host's document, so their rows do not arrive twice", async () => {
        // Every cell here was created during the host's own session, so none of
        // them is keyed the way seeding from a file would key it.
        shared.sheets.find((s) => s.kind !== "cx")!.data = [];
        const sheetId = shared.sheets.find((s) => s.kind !== "cx")!.id;
        const hostSide = side(shared);
        const host = (await startCollabSession({
            createLink: net.create(ALEX),
            roundId: shared.id,
            appVersion: "0.11.0",
            doc: hostSide.doc,
            apply: hostSide.apply,
        }))!;
        hostSide.edit(sheetId, 0, 0, "perm do both");
        hostSide.edit(sheetId, 0, 1, "then the CP");

        const joined = await joinRound({
            ticket: encodeTicket(host.share("partner")),
            createLink: net.create("sam"),
            appVersion: "0.11.0",
            fs,
        });
        const round = parseFlowFile((await fs.readFlow(joined!.path))!.text);
        // The open path's order: the store loads and seeds, then the sidecar
        // upgrades what the seed guessed.
        useFlowStore.getState().loadRound(round);
        await recoverReplica(round, serializeFlow(round));

        // The host sending its whole state again is what a reconnect does.
        applyRemoteDoc(round, hostSide.doc());
        expect(
            useFlowStore
                .getState()
                .round!.sheets.find((s) => s.id === sheetId)!
                .data.map((r) => r[0]),
        ).toEqual(["perm do both", "then the CP"]);
    });

    it("puts the host's first document through the gate every later one goes through", async () => {
        const sheetId = shared.sheets.find((s) => s.kind !== "cx")!.id;
        const hostSide = side(shared);
        // The host chooses every byte of a guest's first document, and a join
        // is the one path that used to skip `merge`. A rank with a trailing
        // zero digit is the shape `rankBetween` cannot subdivide, so a cell
        // carrying one throws on the guest's next insert into that column and
        // comes back from the sidecar after every restart.
        const poisoned: CollabDoc = {
            ...hostSide.doc(),
            sheets: {
                ...hostSide.doc().sheets,
                [sheetId]: {
                    ...hostSide.doc().sheets[sheetId],
                    cells: {
                        ...hostSide.doc().sheets[sheetId].cells,
                        bottom: {
                            col: 0,
                            rank: "zzzzz0",
                            actor: ALEX,
                            text: "theirs",
                            textStamp: { ms: 9_000, counter: 0, actor: ALEX },
                            meta: {},
                            metaStamp: { ms: 9_000, counter: 0, actor: ALEX },
                            deleted: null,
                        },
                    },
                },
            },
        };
        const host = (await startCollabSession({
            createLink: net.create(ALEX),
            roundId: shared.id,
            appVersion: "0.11.0",
            doc: () => poisoned,
            apply: hostSide.apply,
        }))!;

        const joined = await joinRound({
            ticket: encodeTicket(host.share("partner")),
            createLink: net.create("sam"),
            appVersion: "0.11.0",
            fs,
        });
        const round = parseFlowFile((await fs.readFlow(joined!.path))!.text);
        await recoverReplica(round, serializeFlow(round));
        expect(Object.keys(getReplica()!.sheets[sheetId].cells)).not.toContain("bottom");
    });
});
