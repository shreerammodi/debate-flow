import { beforeEach, describe, expect, it } from "vitest";

import { seedDoc } from "@/lib/collab/doc";
import { merge, type DroppedCell } from "@/lib/collab/merge";
import { applyOp, type CollabOp, type OpContext } from "@/lib/collab/ops";
import { createClock } from "@/lib/collab/stamp";
import { attachSync, type PeerSync } from "@/lib/collab/sync";
import type { CollabDoc } from "@/lib/collab/types";
import { makeFlowRound, type FlowRound } from "@/lib/model/flow";

import { createFaultNet } from "./faultNet";
import { makePrng } from "./random";

const net = createFaultNet();

/** Time the test owns, so a repair tick fires when the scenario says it does. */
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
            for (let guard = 0; guard < 100; guard++) {
                const due = pending.filter((p) => p.at <= now);
                if (due.length === 0) return;
                pending = pending.filter((p) => p.at > now);
                for (const p of due) p.fn();
            }
            throw new Error("timers rearmed without end");
        },
        reset() {
            pending = [];
            now = 0;
        },
    };
}

const clock = manualClock();

function sharedRound(): FlowRound {
    const r = makeFlowRound({});
    for (const sheet of r.sheets) {
        sheet.data = [
            ["perm", "link"],
            ["cap bad", "turn"],
            ["extend", "drop"],
        ];
    }
    return r;
}

/**
 * One peer: its replica, its clock, and the loss it has been told about.
 *
 * A real instance is this plus a grid and a file. Neither participates in
 * convergence, so neither is here.
 */
interface Instance {
    id: string;
    doc: CollabDoc;
    ctx: OpContext;
    sync: PeerSync;
    /** Every cell this peer has had to bury, in the order it learned of them. */
    dropped: DroppedCell[];
    edit(op: CollabOp): void;
}

/** Two peers on one round, connected, with nothing yet in flight. */
async function pair(round: FlowRound) {
    const made: Record<string, Instance> = {};

    function build(id: string, startMs: number): Instance {
        let t = startMs;
        const inst: Instance = {
            id,
            doc: seedDoc(round),
            ctx: { actor: id, clock: createClock(id, () => t++) },
            sync: null as unknown as PeerSync,
            dropped: [],
            edit(op) {
                inst.doc = applyOp(inst.doc, op, inst.ctx);
                inst.sync.notifyLocalChange();
            },
        };
        return inst;
    }

    const alex = build("alex", 1_000);
    const sam = build("sam", 5_000);
    made.alex = alex;
    made.sam = sam;

    function wire(self: Instance, conn: Parameters<typeof attachSync>[0]["conn"]): PeerSync {
        return attachSync({
            conn,
            endpointId: self.id,
            // The harness runs alex against sam and nothing else, so the peer on
            // the far side of either link is the other one.
            from: self.id === "alex" ? "sam" : "alex",
            doc: () => self.doc,
            apply: (incoming) => {
                const result = merge(self.doc, incoming);
                self.doc = result.doc;
                self.dropped.push(...result.dropped);
                return result.dropped;
            },
            schedule: clock.schedule,
        });
    }

    const listenerLink = await net.create("sam")({ discovery: "mdns", relay: false });
    let samConn: Parameters<typeof attachSync>[0]["conn"] | null = null;
    await listenerLink.listen((peer) => {
        samConn = peer;
    });
    const diallerLink = await net.create("alex")({ discovery: "mdns", relay: false });
    const alexConn = await diallerLink.dial("sam");

    alex.sync = wire(alex, alexConn);
    sam.sync = wire(sam, samConn!);
    return made as { alex: Instance; sam: Instance };
}

/** Runs every push debounce and every repair tick, then delivers. */
function settle(seconds = 6): void {
    clock.advance(seconds * 1_000);
    net.flush();
    clock.advance(seconds * 1_000);
    net.flush();
}

/**
 * The two replicas agree on every sheet.
 *
 * Round scalars are excluded on purpose: `rfd` is renamed out and dropped in
 * by design, so the two sides are never meant to hold the same notes. Sheets
 * are where convergence is a claim.
 */
function converged(a: Instance, b: Instance): void {
    expect(a.doc.sheets).toEqual(b.doc.sheets);
}

/** Every cell text a peer can currently see, which is what a debater reads. */
function visible(inst: Instance): string[] {
    return Object.values(inst.doc.sheets)
        .flatMap((sheet) => Object.values(sheet.cells))
        .filter((cell) => cell.deleted === null)
        .map((cell) => cell.text)
        .filter((text): text is string => text !== null);
}

beforeEach(() => {
    net.reset();
    clock.reset();
});

describe("a link that dies mid-burst", () => {
    it("converges once the peers are talking again, losing nothing", async () => {
        const { alex, sam } = await pair(sharedRound());
        const sheetId = Object.keys(alex.doc.sheets)[0];

        alex.edit({ kind: "cellText", sheetId, col: 0, row: 0, text: "first" });
        alex.edit({ kind: "cellText", sheetId, col: 0, row: 1, text: "second" });
        clock.advance(50);

        // The burst is written but still on the wire when the link dies.
        expect(net.inFlight()).toBeGreaterThan(0);
        const lost = net.discardInFlight();
        expect(lost).toBeGreaterThan(0);

        // The repair tick states what each side has seen, and the reply is
        // everything above it. Nothing re-sends the lost burst by name.
        settle();
        converged(alex, sam);
        expect(visible(sam)).toContain("first");
        expect(visible(sam)).toContain("second");
        expect(sam.dropped).toEqual([]);
        expect(alex.dropped).toEqual([]);
    });
});

describe("a peer that restarts", () => {
    it("comes back with its file and is caught up by the repair path", async () => {
        const round = sharedRound();
        const { alex, sam } = await pair(round);
        const sheetId = Object.keys(alex.doc.sheets)[0];

        alex.edit({ kind: "cellText", sheetId, col: 1, row: 0, text: "before the crash" });
        settle();
        converged(alex, sam);

        // Sam's process exits. Its replica survives on disk, its link does not.
        const survived = sam.doc;
        sam.sync.stop();
        net.killLinks();

        alex.edit({ kind: "cellText", sheetId, col: 1, row: 1, text: "while sam was gone" });
        clock.advance(1_000);
        net.discardInFlight();

        // Sam restarts: same file, fresh connection.
        net.reset();
        clock.reset();
        const restarted = await pair(round);
        restarted.sam.doc = survived;
        restarted.alex.doc = alex.doc;

        settle();
        converged(restarted.alex, restarted.sam);
        expect(visible(restarted.sam)).toContain("while sam was gone");
        expect(restarted.sam.dropped).toEqual([]);
    });
});

describe("delivery that arrives out of order", () => {
    it("converges on the same document whatever the order", async () => {
        const orders = [1, 2, 3, 4, 5].map((seed) => {
            return { seed, rng: makePrng(seed) };
        });

        const results: string[][] = [];
        for (const { rng } of orders) {
            net.reset();
            clock.reset();
            const { alex, sam } = await pair(sharedRound());
            const sheetId = Object.keys(alex.doc.sheets)[0];

            alex.edit({ kind: "cellText", sheetId, col: 0, row: 0, text: "alex one" });
            sam.edit({ kind: "cellText", sheetId, col: 1, row: 0, text: "sam one" });
            alex.edit({ kind: "insertRow", sheetId, row: 1 });
            sam.edit({ kind: "cellText", sheetId, col: 0, row: 2, text: "sam two" });
            // Both rename the sheet. Scalars merge through a different path
            // than cells do, and only a stamp comparison makes that path
            // order-independent too.
            alex.edit({ kind: "sheetField", sheetId, path: "title", value: "1AC" });
            sam.edit({ kind: "sheetField", sheetId, path: "title", value: "1NC" });
            clock.advance(50);

            for (let i = 0; i < 6; i++) {
                net.flushShuffled(rng);
                clock.advance(6_000);
            }
            net.flush();

            converged(alex, sam);
            results.push([
                ...visible(alex).sort(),
                `title=${String(alex.doc.sheets[sheetId].fields.title.value)}`,
            ]);
        }

        // Every ordering lands on one document, not merely on a self-consistent
        // pair: a merge that depended on arrival order would differ here.
        for (const seen of results) expect(seen).toEqual(results[0]);
    });
});

describe("one side delayed by seconds", () => {
    it("catches up without the fast side having to know it was behind", async () => {
        const { alex, sam } = await pair(sharedRound());
        const sheetId = Object.keys(alex.doc.sheets)[0];

        // Alex works for several repair cycles while nothing is delivered.
        for (let i = 0; i < 5; i++) {
            alex.edit({ kind: "cellText", sheetId, col: 0, row: i % 3, text: `burst ${i}` });
            clock.advance(6_000);
        }
        expect(net.inFlight()).toBeGreaterThan(0);

        settle();
        converged(alex, sam);
        expect(visible(sam)).toContain("burst 4");
        expect(sam.dropped).toEqual([]);
    });
});

describe("a partition that heals", () => {
    it("converges, and reports the write a concurrent delete buried", async () => {
        const { alex, sam } = await pair(sharedRound());
        const sheetId = Object.keys(alex.doc.sheets)[0];

        settle();
        net.partition("alex", "sam");

        // Both sides work on the same row, unable to see each other.
        sam.edit({ kind: "cellText", sheetId, col: 0, row: 2, text: "sam's evidence" });
        alex.edit({ kind: "removeRow", sheetId, row: 2 });
        clock.advance(6_000);

        // Queued, not delivered: a partition holds messages rather than eating
        // them, which is what makes the heal a real test.
        expect(net.inFlight()).toBeGreaterThan(0);

        net.heal();
        settle();
        converged(alex, sam);

        // A delete wins unconditionally, so the text is gone from both sides.
        expect(visible(alex)).not.toContain("sam's evidence");
        expect(visible(sam)).not.toContain("sam's evidence");

        // And it is reported rather than vanishing, which is the whole point
        // of the loss report: this is the one loss a debater cannot see.
        const reported = [...alex.dropped, ...sam.dropped];
        expect(reported.map((d) => d.text)).toContain("sam's evidence");
        expect(reported.find((d) => d.text === "sam's evidence")?.writtenBy).toBe("sam");
    });

    it("reports nothing when the two sides never touched the same row", async () => {
        const { alex, sam } = await pair(sharedRound());
        const sheetId = Object.keys(alex.doc.sheets)[0];

        settle();
        net.partition("alex", "sam");
        alex.edit({ kind: "cellText", sheetId, col: 0, row: 0, text: "alex only" });
        sam.edit({ kind: "cellText", sheetId, col: 1, row: 1, text: "sam only" });
        clock.advance(6_000);

        net.heal();
        settle();
        converged(alex, sam);
        expect(visible(alex)).toContain("alex only");
        expect(visible(alex)).toContain("sam only");
        expect([...alex.dropped, ...sam.dropped]).toEqual([]);
    });
});

describe("a field this build does not read", () => {
    it("survives a round trip through the peer that ignored it", async () => {
        const { alex, sam } = await pair(sharedRound());
        const sheetId = Object.keys(alex.doc.sheets)[0];

        // A newer build wrote a sheet field this one has no code for. Sam
        // carries it without understanding it, and hands it back intact.
        alex.doc.sheets[sheetId].fields["futureField"] = {
            value: "kept",
            stamp: alex.ctx.clock.tick(),
        };
        alex.sync.notifyLocalChange();
        settle();

        expect(sam.doc.sheets[sheetId].fields["futureField"]?.value).toBe("kept");

        // Sam edits something else and ships its own state back.
        sam.edit({ kind: "cellText", sheetId, col: 0, row: 0, text: "sam wrote this" });
        settle();
        converged(alex, sam);
        expect(alex.doc.sheets[sheetId].fields["futureField"]?.value).toBe("kept");
    });
});
