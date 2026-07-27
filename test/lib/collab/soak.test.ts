import { beforeEach, describe, expect, it } from "vitest";

import { vectorOf } from "@/lib/collab/delta";
import { seedDoc } from "@/lib/collab/doc";
import { merge } from "@/lib/collab/merge";
import { applyOp, type CollabOp, type OpContext } from "@/lib/collab/ops";
import { createClock } from "@/lib/collab/stamp";
import { attachSync, type PeerSync } from "@/lib/collab/sync";
import type { CollabDoc } from "@/lib/collab/types";
import { makeFlowRound, makeFlowSheet, type FlowRound } from "@/lib/model/flow";

import { createFaultNet } from "./faultNet";
import { makePrng } from "./random";

/**
 * A round the length of a real one, driven over a misbehaving link for as many
 * writes as two debaters make in an elim.
 *
 * The small fault tests each pin one failure. This asks the question those
 * cannot: after hundreds of writes, reorderings, partitions and drops, do the
 * two machines still show the same flow, and is every word still there.
 */

const net = createFaultNet();

const SHEETS = 4;
const COLS = 7;
const SEEDED_ROWS = 12;
const SEEDS = [1, 7, 42, 101, 2024];

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
            for (let guard = 0; guard < 200; guard++) {
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

function bigRound(): FlowRound {
    const round = makeFlowRound({});
    while (round.sheets.length < SHEETS) {
        round.sheets.push(
            makeFlowSheet({
                title: `${round.sheets.length}.`,
                group: "neg",
                order: round.sheets.length,
            }),
        );
    }
    for (const sheet of round.sheets) {
        sheet.data = Array.from({ length: SEEDED_ROWS }, (_, r) =>
            Array.from({ length: COLS }, (_, c) => (r % 4 === 0 ? null : `seed ${r}.${c}`)),
        );
    }
    return round;
}

interface Instance {
    id: string;
    doc: CollabDoc;
    ctx: OpContext;
    sync: PeerSync;
    /** Every text this peer wrote, so loss can be told from a merge decision. */
    wrote: Set<string>;
    edit(op: CollabOp): void;
}

async function pair(round: FlowRound): Promise<{ alex: Instance; sam: Instance }> {
    function build(id: string, startMs: number): Instance {
        let t = startMs;
        const inst: Instance = {
            id,
            doc: seedDoc(round),
            ctx: { actor: id, clock: createClock(id, () => t++) },
            sync: null as unknown as PeerSync,
            wrote: new Set<string>(),
            edit(op) {
                inst.doc = applyOp(inst.doc, op, inst.ctx);
                if (op.kind === "cellText" && op.text) inst.wrote.add(op.text);
                inst.sync.notifyLocalChange();
            },
        };
        return inst;
    }

    const alex = build("alex", 1_000);
    const sam = build("sam", 5_000);

    const listenerLink = await net.create("sam")({ discovery: "mdns", relay: false });
    let samConn: Parameters<typeof attachSync>[0]["conn"] | null = null;
    await listenerLink.listen((peer) => {
        samConn = peer;
    });
    const diallerLink = await net.create("alex")({ discovery: "mdns", relay: false });
    const alexConn = await diallerLink.dial("sam");

    const wire = (self: Instance, conn: Parameters<typeof attachSync>[0]["conn"]): PeerSync =>
        attachSync({
            conn,
            endpointId: self.id,
            doc: () => self.doc,
            apply: (incoming) => {
                const result = merge(self.doc, incoming);
                self.doc = result.doc;
                // What `replaceReplicaDoc` does on the real inbound path. The
                // two clocks start hours apart here, and without this the peer
                // that is behind stamps its next write below what it just
                // received and quietly overwrites it with an older value.
                for (const stamp of Object.values(vectorOf(incoming)))
                    self.ctx.clock.observe(stamp);
                return result.dropped;
            },
            schedule: clock.schedule,
        });

    alex.sync = wire(alex, alexConn);
    sam.sync = wire(sam, samConn!);
    return { alex, sam };
}

/** Runs every debounce and repair tick until the link is quiet. */
function quiesce(): void {
    for (let i = 0; i < 6; i++) {
        clock.advance(6_000);
        net.flush();
    }
}

function liveTexts(inst: Instance): Set<string> {
    const out = new Set<string>();
    for (const sheet of Object.values(inst.doc.sheets)) {
        for (const cell of Object.values(sheet.cells)) {
            if (cell.deleted === null && cell.text) out.add(cell.text);
        }
    }
    return out;
}

beforeEach(() => {
    net.reset();
    clock.reset();
});

describe("a full round over a link that misbehaves", () => {
    it.each(SEEDS)("converges on the same flow (seed %i)", async (seed) => {
        const rng = makePrng(seed);
        const round = bigRound();
        const { alex, sam } = await pair(round);
        const sheetIds = Object.keys(alex.doc.sheets);
        let partitioned = false;

        for (let step = 0; step < 400; step++) {
            const who = rng() < 0.5 ? alex : sam;
            const sheetId = sheetIds[Math.floor(rng() * sheetIds.length)];
            const col = Math.floor(rng() * COLS);
            const row = Math.floor(rng() * (SEEDED_ROWS + 8));
            const roll = rng();
            if (roll < 0.7) {
                who.edit({ kind: "cellText", sheetId, col, row, text: `${who.id} ${step}` });
            } else if (roll < 0.8) {
                who.edit({ kind: "cellMeta", sheetId, col, row, meta: { bold: rng() < 0.5 } });
            } else if (roll < 0.9) {
                who.edit({ kind: "insertCell", sheetId, col, row });
            } else {
                who.edit({ kind: "removeRow", sheetId, row });
            }

            // What a tournament network does to a QUIC stream: reorders
            // against the other direction, and stalls. It does not silently
            // drop one message and carry on - a stream that loses anything
            // fails, and a failure rebuilds this sync from a full state, which
            // the reconnect test below covers on its own.
            const event = rng();
            if (event < 0.4) {
                clock.advance(40);
                net.flushShuffled(rng);
            } else if (event < 0.5 && !partitioned) {
                net.partition("alex", "sam");
                partitioned = true;
            } else if (event < 0.6 && partitioned) {
                net.heal();
                partitioned = false;
                clock.advance(40);
                net.flush();
            }
        }

        if (partitioned) net.heal();
        quiesce();

        expect(net.inFlight()).toBe(0);
        expect(alex.doc.sheets).toEqual(sam.doc.sheets);
    });

    it.each(SEEDS)("loses no word that nobody deleted (seed %i)", async (seed) => {
        const rng = makePrng(seed * 13 + 1);
        const round = bigRound();
        const { alex, sam } = await pair(round);
        const sheetIds = Object.keys(alex.doc.sheets);

        // Writes only, and every one to a cell of its own: with no delete and
        // no overwrite in play, every word typed must survive on both machines
        // whatever the link did in between.
        for (let step = 0; step < 300; step++) {
            const who = rng() < 0.5 ? alex : sam;
            const sheetId = sheetIds[Math.floor(rng() * sheetIds.length)];
            // Each peer owns its own columns, so nothing is a genuine conflict
            // and last-writer-wins is never entitled to drop a word.
            const col = who === alex ? 0 : 4;
            who.edit({ kind: "cellText", sheetId, col, row: step, text: `${who.id} ${step}` });

            if (rng() < 0.5) {
                clock.advance(40);
                net.flushShuffled(rng);
            }
        }
        quiesce();

        const onAlex = liveTexts(alex);
        const onSam = liveTexts(sam);
        for (const text of [...alex.wrote, ...sam.wrote]) {
            expect(onAlex.has(text), `${text} is gone from alex`).toBe(true);
            expect(onSam.has(text), `${text} is gone from sam`).toBe(true);
        }
        expect(alex.doc.sheets).toEqual(sam.doc.sheets);
    });

    it("catches a peer up on everything it missed while its link was down", async () => {
        const round = bigRound();
        const { alex, sam } = await pair(round);
        const sheetId = Object.keys(alex.doc.sheets)[1];

        for (let i = 0; i < 40; i++) {
            alex.edit({ kind: "cellText", sheetId, col: 0, row: i, text: `alex ${i}` });
        }
        quiesce();
        expect(alex.doc.sheets).toEqual(sam.doc.sheets);

        // The link goes down and both sides keep flowing into it.
        net.partition("alex", "sam");
        for (let i = 0; i < 40; i++) {
            sam.edit({ kind: "cellText", sheetId, col: 4, row: i, text: `sam ${i}` });
            alex.edit({ kind: "cellText", sheetId, col: 1, row: i, text: `alex late ${i}` });
        }
        clock.advance(40);
        net.heal();
        quiesce();

        expect(alex.doc.sheets).toEqual(sam.doc.sheets);
        const visible = liveTexts(alex);
        for (let i = 0; i < 40; i++) {
            expect(visible.has(`alex ${i}`)).toBe(true);
            expect(visible.has(`alex late ${i}`)).toBe(true);
            expect(visible.has(`sam ${i}`)).toBe(true);
        }
    });

    /**
     * Two laptops whose clocks disagree, both typing the same cells.
     *
     * A stamp leads with wall time, so the machine that is behind stamps its
     * next write below the one it just merged, loses last-writer-wins on the
     * far side, and keeps its own text locally: two flows that disagree and
     * never exchange another word about it. The local clock taking on every
     * stamp it receives is the whole defence, and nothing else in the system
     * notices if it stops happening.
     */
    it("keeps the machine whose clock runs behind from losing what it types", async () => {
        const round = bigRound();
        const { alex, sam } = await pair(round);
        const sheetId = Object.keys(alex.doc.sheets)[1];

        // Sam is an hour ahead. Alex writes last, every time.
        for (let i = 0; i < 30; i++) {
            sam.edit({ kind: "cellText", sheetId, col: 0, row: i, text: `sam ${i}` });
            clock.advance(40);
            net.flush();
            alex.edit({ kind: "cellText", sheetId, col: 0, row: i, text: `alex wins ${i}` });
            clock.advance(40);
            net.flush();
        }
        quiesce();

        expect(alex.doc.sheets).toEqual(sam.doc.sheets);
        const onSam = liveTexts(sam);
        for (let i = 0; i < 30; i++) {
            expect(onSam.has(`alex wins ${i}`), `alex wins ${i} never reached sam`).toBe(true);
        }
    });
});
