import { beforeEach, describe, expect, it } from "vitest";

import { projectDoc, seedDoc } from "@/lib/collab/doc";
import { merge } from "@/lib/collab/merge";
import { applyOp, type CollabOp, type OpContext } from "@/lib/collab/ops";
import type { PeerConn } from "@/lib/collab/peerLink";
import { createMemoryNet } from "@/lib/collab/peerLinkMemory";
import { startCollabSession } from "@/lib/collab/session";
import { createClock } from "@/lib/collab/stamp";
import type { CollabDoc } from "@/lib/collab/types";
import { makeFlowRound, type FlowRound } from "@/lib/model/flow";
import { useFlowStore } from "@/lib/store/useFlowStore";

const net = createMemoryNet();

/**
 * One peer: a replica, a clock, and the links it holds. Phase 0 has no sync
 * loop yet, so this stands in for one: every local edit ships the whole
 * replica, and every arrival merges.
 */
class Peer {
    doc: CollabDoc;
    readonly ctx: OpContext;
    private links: PeerConn[] = [];

    constructor(
        readonly actor: string,
        round: FlowRound,
        startMs: number,
    ) {
        let t = startMs;
        this.doc = seedDoc(round);
        this.ctx = { actor, clock: createClock(actor, () => t++) };
    }

    attach(conn: PeerConn): void {
        this.links.push(conn);
        conn.onMessage((msg) => {
            if (msg.type !== "state" && msg.type !== "delta") return;
            this.doc = merge(this.doc, msg.doc).doc;
        });
    }

    edit(op: CollabOp): void {
        this.doc = applyOp(this.doc, op, this.ctx);
        for (const conn of this.links) conn.send({ type: "delta", doc: this.doc });
    }
}

beforeEach(() => {
    net.reset();
    useFlowStore.setState({ collabEnabled: true, collabRelayEnabled: true });
});

describe("two peers over the memory transport", () => {
    it("converges on the same flow after editing at once", async () => {
        const round = makeFlowRound({});
        const sheetId = round.sheets.find((s) => s.kind !== "cx")!.id;
        round.sheets.find((s) => s.kind !== "cx")!.data = [
            ["perm do both", "no link"],
            ["cap bad", "turn"],
        ];

        const alex = new Peer("alex", round, 1_000);
        const sam = new Peer("sam", round, 1_000);

        const host = await startCollabSession({
            createLink: net.create("alex"),
            onPeer: (conn) => alex.attach(conn),
        });
        const guest = await startCollabSession({
            createLink: net.create("sam"),
            peers: ["alex"],
        });
        sam.attach(guest!.peers[0]);

        // Neither peer touches a cell the other is in, which is the round the
        // feature is designed for.
        alex.edit({ kind: "cellText", sheetId, col: 0, row: 0, text: "perm do both, then CP" });
        sam.edit({ kind: "cellText", sheetId, col: 1, row: 1, text: "turn, ext. Smith" });
        sam.edit({ kind: "insertRow", sheetId, row: 1 });
        alex.edit({ kind: "roundField", path: "scouting.tournament", value: "Harvard" });

        expect(alex.doc).toEqual(sam.doc);
        const mine = projectDoc(alex.doc, round);
        const theirs = projectDoc(sam.doc, round);
        expect(mine).toEqual(theirs);
        expect(mine.scouting.tournament).toBe("Harvard");
        expect(mine.sheets.find((s) => s.id === sheetId)!.data).toEqual([
            ["perm do both, then CP", "no link"],
            [null, null],
            ["cap bad", "turn, ext. Smith"],
        ]);

        await host!.stop();
        await guest!.stop();
    });

    it("heals a partition, and names the cells the delete buried", async () => {
        const round = makeFlowRound({});
        const flow = round.sheets.find((s) => s.kind !== "cx")!;
        const sheetId = flow.id;
        flow.data = [
            ["perm", "link"],
            ["cap bad", "turn"],
        ];

        const alex = new Peer("alex", round, 1_000);
        const sam = new Peer("sam", round, 5_000);

        // Partitioned: both edit with no link between them.
        alex.edit({ kind: "removeRow", sheetId, row: 0 });
        sam.edit({ kind: "cellText", sheetId, col: 0, row: 0, text: "perm is severance" });

        const healed = merge(alex.doc, sam.doc);
        expect(healed.dropped).toEqual([]);
        const fromSam = merge(sam.doc, alex.doc);
        expect(fromSam.doc).toEqual(healed.doc);
        // Sam held the text alive, so Sam is the peer that is told.
        expect(fromSam.dropped.map((d) => [d.text, d.deletedBy])).toEqual([
            ["perm is severance", "alex"],
            ["link", "alex"],
        ]);
        expect(projectDoc(healed.doc, round).sheets.find((s) => s.id === sheetId)!.data).toEqual([
            ["cap bad", "turn"],
        ]);
    });
});
