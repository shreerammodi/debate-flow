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

import { vectorOf } from "./delta";
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

/**
 * Takes on this machine's own identity for everything written from here.
 *
 * Solo, the actor is empty: two peers who open the same file derive byte
 * identical replicas, which is what lets them sync with no negotiation. That
 * only holds for cells the file already had. A cell created during a session
 * needs an author, because a cell's identity is its column, its rank, and its
 * creator, and two peers inserting at one position derive the same rank. With
 * one shared identity those two cells collide on the same key and one is lost
 * with nothing reporting it.
 *
 * The document carries over untouched; only later writes are affected.
 */
export function adoptReplicaActor(actor: string): void {
    if (!live || live.actor === actor) return;
    live = { doc: live.doc, clock: createClock(actor), actor };
}

export function clearReplica(): void {
    live = null;
}

/**
 * Told after every local write, so a live session can push it out. A bridge
 * rather than a direct call because the runtime owns the session and already
 * reads this module; the dependency only ever runs one way.
 */
let onLocalChange: (() => void) | null = null;

export function setLocalChangeListener(fn: (() => void) | null): void {
    onLocalChange = fn;
}

/**
 * Takes the document a merge produced. The clock and the identity survive: a
 * fresh clock could repeat a stamp this machine has already written inside the
 * same millisecond, and last-writer-wins would then keep whichever of the two
 * writes it saw first.
 *
 * The incoming stamps raise the clock, so a peer whose wall clock runs ahead
 * cannot win every later tie by sitting in the future.
 */
export function replaceReplicaDoc(doc: CollabDoc): void {
    if (!live) return;
    for (const stamp of Object.values(vectorOf(doc))) live.clock.observe(stamp);
    live = { ...live, doc };
}

export function getReplica(): CollabDoc | null {
    return live?.doc ?? null;
}

/** This machine's actor in the live replica, or "" when nothing is open. */
export function replicaActor(): string {
    return live?.actor ?? "";
}

export function replicaRoundId(): string | null {
    return live?.doc.roundId ?? null;
}

/** Mirrors one local edit. A no-op with no round open, so a late hook is safe. */
export function recordOp(op: CollabOp): void {
    if (!live) return;
    live.doc = applyOp(live.doc, op, { actor: live.actor, clock: live.clock });
    onLocalChange?.();
}

/**
 * Re-derives one sheet from the store's copy.
 *
 * A block move no longer comes through here: it emits ops a peer can apply.
 *
 * ponytail: an insert-paste and a CardMirror send still re-seed, because both
 * rearrange a column from outside the op path and neither happens while a
 * partner is typing into the same sheet. Re-seeding re-keys every cell from
 * its row position, so a peer holding the old keys disagrees; express these
 * two as ops before either can run mid-session.
 */
export function resyncSheet(sheet: FlowSheet): void {
    if (!live) return;
    live.doc = {
        ...live.doc,
        sheets: { ...live.doc.sheets, [sheet.id]: seedSheet(sheet, ORIGIN_STAMP) },
    };
    onLocalChange?.();
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
