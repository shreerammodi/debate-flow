/**
 * Shadow mode: a real round runs shared while the partner's writes are logged
 * and diffed rather than trusted.
 *
 * Remote changes land in a shadow replica and never on the live one, so the
 * round in progress is never at risk. Diffing the two projections says what
 * the merge would have done to the grid, in the terms a debater reads: a
 * sheet, a cell, the text there now and the text that would have replaced it.
 * That is how the feature meets real rounds before a round depends on it.
 *
 * The shadow keeps absorbing this machine's own writes, so the only thing
 * standing between the two replicas is the partner. A log a reader learns to
 * skim is worth nothing, and one false alarm is enough to teach that.
 *
 * Every observation records an entry, including one whose merge changes
 * nothing. A run of empty entries is the reading that earns confidence: the
 * messages arrived and none of them would have touched this flow. Dropping
 * those entries would make a quiet link and a correct one look identical.
 */

import { sortedSheets, type FlowRound, type FlowSheet } from "@/lib/model/flow";

import { projectDoc } from "./doc";
import { merge, type DroppedCell } from "./merge";
import type { CollabDoc } from "./types";

/** One cell a remote change would have altered, at the level a debater sees. */
export interface ShadowDiff {
    sheetId: string;
    col: number;
    row: number;
    /** What is on this machine's grid now. Empty string for a blank cell. */
    mine: string;
    /** What the merge would have put there. */
    theirs: string;
}

export interface ShadowEntry {
    /** Epoch ms. */
    at: number;
    /** The peer whose change this was. */
    from: string;
    diffs: ShadowDiff[];
    /** Cells the merge would have buried. */
    dropped: DroppedCell[];
}

export interface Shadow {
    /**
     * Merges a remote document into the shadow replica and reports what it
     * would have changed on the live grid. Never touches the live document.
     */
    observe(from: string, incoming: CollabDoc): ShadowEntry;
    entries(): readonly ShadowEntry[];
    clear(): void;
}

export interface ShadowDeps {
    /** The live replica, read fresh on every observation. */
    doc(): CollabDoc;
    /** The round the projection is layered onto. */
    base(): FlowRound;
    now?(): number;
}

function diffSheet(
    sheetId: string,
    mine: FlowSheet | undefined,
    theirs: FlowSheet | undefined,
    out: ShadowDiff[],
): void {
    const rows = Math.max(mine?.data.length ?? 0, theirs?.data.length ?? 0);
    for (let row = 0; row < rows; row++) {
        const mineRow = mine?.data[row];
        const theirsRow = theirs?.data[row];
        const cols = Math.max(mineRow?.length ?? 0, theirsRow?.length ?? 0);
        for (let col = 0; col < cols; col++) {
            // A blank cell and a cell off the end of the grid read the same way.
            const here = mineRow?.[col] ?? "";
            const there = theirsRow?.[col] ?? "";
            if (here !== there) out.push({ sheetId, col, row, mine: here, theirs: there });
        }
    }
}

function diffRounds(mine: FlowRound, theirs: FlowRound): ShadowDiff[] {
    const unvisited = new Map(sortedSheets(mine).map((s) => [s.id, s]));
    const diffs: ShadowDiff[] = [];
    for (const sheet of sortedSheets(theirs)) {
        diffSheet(sheet.id, unvisited.get(sheet.id), sheet, diffs);
        unvisited.delete(sheet.id);
    }
    // A sheet only this machine still holds: the merge would have taken it away.
    for (const sheet of unvisited.values()) diffSheet(sheet.id, sheet, undefined, diffs);
    return diffs;
}

export function createShadow(deps: ShadowDeps): Shadow {
    const now = deps.now ?? Date.now;
    // Null until the first observation, because a session can exist before a
    // round is loaded, and again after clear so the next one re-bases.
    let replica: CollabDoc | null = null;
    let log: ShadowEntry[] = [];

    return {
        observe(from, incoming) {
            const live = deps.doc();
            // The shadow takes this machine's own writes back before the
            // remote one lands, so a delta that carries only the partner's
            // cells never reads as the partner blanking a cell the host just
            // typed. Only the remote merge's losses are the partner's doing,
            // so this merge's report is not the one the entry carries.
            const rebased = merge(replica ?? live, live).doc;
            const merged = merge(rebased, incoming);
            replica = merged.doc;
            const base = deps.base();
            const entry: ShadowEntry = {
                at: now(),
                from,
                diffs: diffRounds(projectDoc(live, base), projectDoc(replica, base)),
                dropped: merged.dropped,
            };
            log.push(entry);
            return entry;
        },
        entries() {
            return log;
        },
        clear() {
            log = [];
            replica = null;
        },
    };
}
