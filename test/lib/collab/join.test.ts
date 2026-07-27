import { beforeEach, describe, expect, it } from "vitest";

import { seedDoc } from "@/lib/collab/doc";
import { joinRound } from "@/lib/collab/join";
import { merge } from "@/lib/collab/merge";
import { applyOp, type OpContext } from "@/lib/collab/ops";
import { createMemoryNet } from "@/lib/collab/peerLinkMemory";
import { recoverReplica } from "@/lib/collab/persist";
import { forgetRoundPeers, knownRoundPeers } from "@/lib/collab/roundPeers";
import { applyRemoteDoc } from "@/lib/collab/runtime";
import { startCollabSession } from "@/lib/collab/session";
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

let shared: FlowRound;
let fs: FlowFs;

function side(base: FlowRound) {
    let doc = seedDoc(base);
    let t = 1_000;
    const ctx: OpContext = { actor: "alex", clock: createClock("alex", () => t++) };
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
        createLink: net.create("alex"),
        roundId: shared.id,
        appVersion: "0.11.0",
        ...side(shared),
    }))!;
    return encodeTicket(host.share("partner"));
}

beforeEach(async () => {
    net.reset();
    forgetRoundPeers();
    useFlowStore.setState({ collabEnabled: true, collabRelayEnabled: true });
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
        expect(result!.hostEndpointId).toBe("alex");

        const written = await fs.readFlow(result!.path);
        expect(written).not.toBeNull();
        expect(written!.text).toContain("perm");
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

    it("reaches no network at all while shared editing is off", async () => {
        useFlowStore.setState({ collabEnabled: false });
        const result = await joinRound({
            ticket: encodeTicket({
                endpointId: "alex",
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
            createLink: net.create("alex"),
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
            invite: { endpointId: "alex", roundId: shared.id },
            createLink: net.create("sam"),
            appVersion: "0.11.0",
            fs,
        });
        expect(result!.created).toBe(true);
        expect(result!.hostEndpointId).toBe("alex");
        expect((await fs.readFlow(result!.path))!.text).toContain("perm");
    });

    it("is refused by a host who never invited this peer", async () => {
        await startCollabSession({
            createLink: net.create("alex"),
            roundId: shared.id,
            appVersion: "0.11.0",
            ...side(shared),
        });
        await expect(
            joinRound({
                invite: { endpointId: "alex", roundId: shared.id },
                createLink: net.create("sam"),
                appVersion: "0.11.0",
                fs,
            }),
        ).rejects.toThrow();
    });

    it("reaches no network at all while shared editing is off", async () => {
        useFlowStore.setState({ collabEnabled: false });
        const result = await joinRound({
            invite: { endpointId: "alex", roundId: shared.id },
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
        expect(knownRoundPeers(result!.roundId)).toEqual(["alex"]);
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
        expect(await recoverReplica(round, serializeFlow(round))).toEqual(["alex"]);
    });

    it("adopts the host's document, so their rows do not arrive twice", async () => {
        // Every cell here was created during the host's own session, so none of
        // them is keyed the way seeding from a file would key it.
        shared.sheets.find((s) => s.kind !== "cx")!.data = [];
        const sheetId = shared.sheets.find((s) => s.kind !== "cx")!.id;
        const hostSide = side(shared);
        const host = (await startCollabSession({
            createLink: net.create("alex"),
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
});
