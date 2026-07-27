import { seedDoc } from "@/lib/collab/doc";
import type { CollabOp, OpContext } from "@/lib/collab/ops";
import { createClock } from "@/lib/collab/stamp";
import type { CollabDoc } from "@/lib/collab/types";
import { makeFlowRound, type FlowRound } from "@/lib/model/flow";

/** mulberry32: a seeded PRNG, so a failing case replays from its seed alone. */
export function makePrng(seed: number): () => number {
    let a = seed >>> 0;
    return () => {
        a = (a + 0x6d2b79f5) >>> 0;
        let t = Math.imul(a ^ (a >>> 15), 1 | a);
        t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
        return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
    };
}

export interface Replica {
    actor: string;
    ctx: OpContext;
    doc: CollabDoc;
}

/** A round with two sheets and a small grid, the shape a real flow starts as. */
export function sharedRound(): FlowRound {
    const round = makeFlowRound({});
    for (const sheet of round.sheets) {
        sheet.data = [
            ["perm", "link"],
            ["cap bad", "turn"],
            ["extend", null],
        ];
        sheet.meta = { "0,0": { bold: true } };
    }
    return round;
}

export function makeReplica(actor: string, round: FlowRound, startMs: number): Replica {
    let t = startMs;
    return {
        actor,
        ctx: { actor, clock: createClock(actor, () => t++) },
        doc: seedDoc(round),
    };
}

const WORDS = ["perm", "turn", "link", "extend", "drop", "cap", "framing", ""];

export function randomOp(rng: () => number, sheetIds: string[]): CollabOp {
    const sheetId = sheetIds[Math.floor(rng() * sheetIds.length)];
    const col = Math.floor(rng() * 2);
    const row = Math.floor(rng() * 4);
    switch (Math.floor(rng() * 10)) {
        case 0:
        case 1:
        case 2:
        case 3:
            return {
                kind: "cellText",
                sheetId,
                col,
                row,
                text: WORDS[Math.floor(rng() * WORDS.length)],
            };
        case 4:
            return { kind: "cellMeta", sheetId, col, row, meta: { bold: rng() < 0.5 } };
        case 5:
            return { kind: "insertCell", sheetId, col, row };
        case 6:
            return { kind: "removeCell", sheetId, col, row };
        case 7:
            return { kind: "insertRow", sheetId, row };
        case 8:
            return { kind: "removeRow", sheetId, row };
        default:
            return {
                kind: "sheetField",
                sheetId,
                path: "title",
                value: `T${Math.floor(rng() * 5)}`,
            };
    }
}
