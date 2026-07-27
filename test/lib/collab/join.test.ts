import { beforeEach, describe, expect, it } from "vitest";

import { seedDoc } from "@/lib/collab/doc";
import { joinRound } from "@/lib/collab/join";
import { merge } from "@/lib/collab/merge";
import { createMemoryNet } from "@/lib/collab/peerLinkMemory";
import { startCollabSession } from "@/lib/collab/session";
import { encodeTicket } from "@/lib/collab/ticket";
import type { CollabDoc } from "@/lib/collab/types";
import { makeFlowRound, type FlowRound } from "@/lib/model/flow";
import { serializeFlow } from "@/lib/persistence/flowFile";
import type { FlowFs } from "@/lib/persistence/flowFs";
import { createFlowFs } from "@/lib/persistence/flowFsMemory";
import { saveRecents } from "@/lib/persistence/recents";
import { useFlowStore } from "@/lib/store/useFlowStore";

const net = createMemoryNet();

let shared: FlowRound;
let fs: FlowFs;

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
