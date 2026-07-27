/**
 * The replica that runs beside the grid and the store.
 *
 * Module state rather than store state: a keystroke updates this on every
 * change, and routing that through Zustand would re-render the app for a value
 * no component reads. `hotInstance.ts` holds the active grid the same way and
 * for the same reason.
 *
 * It is maintained whether or not a peer exists, so there is one code path and
 * not two. With no peers the only cost is rank bookkeeping. Nothing here may
 * import the store: the store imports this, and the reverse would be a cycle
 * through Zustand's module initialization.
 */

import type { FlowRound, FlowSheet } from "@/lib/model/flow";

import { projectSheet, seedDoc, seedSheet } from "./doc";
import { sheetDigest } from "./hash";
import { applyOp, type CollabOp } from "./ops";
import { createClock, ORIGIN_STAMP, type Clock } from "./stamp";
import type { CollabDoc, CollabSheet } from "./types";

interface Live {
    doc: CollabDoc;
    clock: Clock;
    actor: string;
}

let live: Live | null = null;

/**
 * Opens a round. `seeded` is a document recovered from the sidecar; without one
 * the replica derives itself from the file, which two peers do identically.
 * Always replaces what came before: switching straight from one flow to another
 * never closes the first.
 */
export function seedReplica(round: FlowRound, actor = "", seeded: CollabDoc | null = null): void {
    live = { doc: seeded ?? seedDoc(round), clock: createClock(actor), actor };
}

export function clearReplica(): void {
    live = null;
}

export function getReplica(): CollabDoc | null {
    return live?.doc ?? null;
}

export function replicaRoundId(): string | null {
    return live?.doc.roundId ?? null;
}

/** Mirrors one local edit. A no-op with no round open, so a late hook is safe. */
export function recordOp(op: CollabOp): void {
    if (!live) return;
    live.doc = applyOp(live.doc, op, { actor: live.actor, clock: live.clock });
}

/**
 * Re-derives one sheet from the store's copy.
 *
 * ponytail: a block move, an insert-paste, and a CardMirror send re-seed the
 * sheet rather than describing themselves as ops, because the op union has no
 * move-shaped member and inventing one now would be speculative. Exact while no
 * peer exists. Phase 3 must replace this before a live session may move a
 * block: re-seeding re-derives ranks from row position, which a peer holding
 * the old ranks would not agree with.
 */
export function resyncSheet(sheet: FlowSheet): void {
    if (!live) return;
    live.doc = {
        ...live.doc,
        sheets: { ...live.doc.sheets, [sheet.id]: seedSheet(sheet, ORIGIN_STAMP) },
    };
}

function digestOf(sheet: CollabSheet): string {
    const projected = projectSheet(sheet);
    return sheetDigest(projected.data, projected.meta);
}

/**
 * The sheets whose replica no longer matches the store. A hook that never fired
 * is otherwise a silent divergence; this turns it into a recoverable one.
 */
export function driftedSheetIds(round: FlowRound): string[] {
    if (!live) return [];
    const drifted: string[] = [];
    for (const sheet of round.sheets) {
        const mine = live.doc.sheets[sheet.id];
        if (!mine || mine.deleted !== null) {
            drifted.push(sheet.id);
            continue;
        }
        if (digestOf(mine) !== sheetDigest(sheet.data, sheet.meta)) drifted.push(sheet.id);
    }
    return drifted;
}

/** Repairs every drifted sheet, and names the ones it touched. */
export function healReplica(round: FlowRound): string[] {
    const drifted = driftedSheetIds(round);
    for (const id of drifted) {
        const sheet = round.sheets.find((s) => s.id === id);
        if (sheet) resyncSheet(sheet);
    }
    return drifted;
}
