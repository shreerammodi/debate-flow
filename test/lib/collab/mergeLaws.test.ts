import { describe, expect, it } from "vitest";

import { seedDoc } from "@/lib/collab/doc";
import { merge } from "@/lib/collab/merge";
import { applyOp } from "@/lib/collab/ops";
import { compareStamps, ORIGIN_STAMP, type Stamp } from "@/lib/collab/stamp";
import type { CollabDoc } from "@/lib/collab/types";

import { makePrng, makeReplica, randomOp, sharedRound, type Replica } from "./random";

const SEEDS = [1, 2, 3, 5, 8, 13, 21, 34, 55, 89, 144, 233];

/** Every replica's state after each of its own ops, which is what a peer ships. */
function play(seed: number, actors: string[]): { snapshots: CollabDoc[][]; replicas: Replica[] } {
    const rng = makePrng(seed);
    const round = sharedRound();
    const sheetIds = round.sheets.map((s) => s.id);
    const replicas = actors.map((a, i) => makeReplica(a, round, 1_000 + i * 7));
    const snapshots: CollabDoc[][] = replicas.map(() => []);
    for (let step = 0; step < 24; step++) {
        const which = Math.floor(rng() * replicas.length);
        const r = replicas[which];
        r.doc = applyOp(r.doc, randomOp(rng, sheetIds), r.ctx);
        snapshots[which].push(r.doc);
    }
    return { snapshots, replicas };
}

function shuffled<T>(rng: () => number, items: T[]): T[] {
    const out = items.slice();
    for (let i = out.length - 1; i > 0; i--) {
        const j = Math.floor(rng() * (i + 1));
        [out[i], out[j]] = [out[j], out[i]];
    }
    return out;
}

describe("merge laws", () => {
    it.each(SEEDS)("is commutative (seed %i)", (seed) => {
        const { replicas } = play(seed, ["alex", "sam"]);
        const [a, b] = replicas;
        expect(merge(a.doc, b.doc).doc).toEqual(merge(b.doc, a.doc).doc);
    });

    it.each(SEEDS)("is associative (seed %i)", (seed) => {
        const { replicas } = play(seed, ["alex", "sam", "kim"]);
        const [a, b, c] = replicas;
        const left = merge(merge(a.doc, b.doc).doc, c.doc).doc;
        const right = merge(a.doc, merge(b.doc, c.doc).doc).doc;
        expect(left).toEqual(right);
    });

    it.each(SEEDS)("is idempotent (seed %i)", (seed) => {
        const { replicas } = play(seed, ["alex", "sam"]);
        for (const r of replicas) {
            expect(merge(r.doc, r.doc).doc).toEqual(r.doc);
            expect(merge(r.doc, r.doc).dropped).toEqual([]);
        }
    });

    it.each(SEEDS)("converges under any delivery order (seed %i)", (seed) => {
        const { snapshots, replicas } = play(seed, ["alex", "sam", "kim"]);
        const rng = makePrng(seed * 31 + 7);
        const finals = replicas.map((r, self) => {
            let doc = r.doc;
            const inbox = snapshots.flatMap((list, from) => (from === self ? [] : list));
            for (const incoming of shuffled(rng, inbox)) doc = merge(doc, incoming).doc;
            return doc;
        });
        for (const doc of finals) expect(doc).toEqual(finals[0]);
    });

    it.each(SEEDS)("loses no non-empty cell without reporting it (seed %i)", (seed) => {
        const { snapshots, replicas } = play(seed, ["alex", "sam", "kim"]);
        const rng = makePrng(seed * 17 + 3);
        for (let self = 0; self < replicas.length; self++) {
            let doc = replicas[self].doc;
            const held: { key: string; sheetId: string; text: string; stamp: Stamp }[] = [];
            for (const [sheetId, sheet] of Object.entries(doc.sheets)) {
                for (const [key, cell] of Object.entries(sheet.cells)) {
                    if (cell.deleted === null && (cell.text ?? "").trim() !== "") {
                        held.push({
                            key,
                            sheetId,
                            text: cell.text as string,
                            stamp: cell.textStamp,
                        });
                    }
                }
            }
            const reported = new Set<string>();
            const inbox = snapshots.flatMap((list, from) => (from === self ? [] : list));
            for (const incoming of shuffled(rng, inbox)) {
                const result = merge(doc, incoming);
                doc = result.doc;
                for (const d of result.dropped) reported.add(`${d.sheetId}|${d.col}|${d.rank}`);
            }
            for (const { key, sheetId, text, stamp } of held) {
                const cell = doc.sheets[sheetId].cells[key];
                expect(cell).toBeDefined();
                // A strictly later write replaced the text, which is LWW the
                // debater can see happen and not a loss the merge must report.
                if (cell.text !== text) {
                    expect(compareStamps(stamp, cell.textStamp)).toBeLessThan(0);
                    continue;
                }
                // This replica's own text, still the winner, and now buried.
                if (cell.deleted !== null) {
                    expect(reported.has(`${sheetId}|${cell.col}|${cell.rank}`)).toBe(true);
                }
            }
        }
    });
});

describe("unknown fields", () => {
    it("survives a round trip through a peer that does not read them", () => {
        const round = sharedRound();
        const sheetId = round.sheets[0].id;
        const stamp: Stamp = { ms: 5_000, counter: 0, actor: "sam" };

        const newBuild = seedDoc(round);
        newBuild.round["scouting.decision.peerNotes.sam"] = { value: "voted aff on turns", stamp };
        newBuild.sheets[sheetId].fields.futureField = { value: 7, stamp };
        const someCell = Object.keys(newBuild.sheets[sheetId].cells)[0];
        newBuild.sheets[sheetId].cells[someCell] = {
            ...newBuild.sheets[sheetId].cells[someCell],
            meta: { futureMeta: "keep" },
            metaStamp: stamp,
        };

        // The old build knows none of the three and never writes them.
        const oldBuild = seedDoc(round);
        expect(oldBuild.round["scouting.decision.peerNotes.sam"]).toBeUndefined();

        const echoed = merge(oldBuild, newBuild).doc;
        const back = merge(seedDoc(round), echoed).doc;
        expect(back.round["scouting.decision.peerNotes.sam"].value).toBe("voted aff on turns");
        expect(back.sheets[sheetId].fields.futureField.value).toBe(7);
        expect(back.sheets[sheetId].cells[someCell].meta).toEqual({ futureMeta: "keep" });
        expect(back.round.event.stamp).toEqual(ORIGIN_STAMP);
    });
});
