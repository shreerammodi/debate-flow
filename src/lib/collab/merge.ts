/**
 * The one sync primitive.
 *
 * A delta and a full state message are the same shape, so both arrive here.
 * The result depends on the pair and not on the order they arrived in, which
 * is what lets a dropped link, a restart, and a replayed message all be
 * harmless.
 *
 * A delete wins unconditionally. Resurrecting a cell would re-insert one
 * column's entry and leave that column offset by one row below the deletion
 * point, which reads as a bug and is silent. The cells a delete discards are
 * reported instead, because that is the one loss a debater cannot see.
 */

import { compareStamps, type Stamp } from "./stamp";
import type { CollabCell, CollabDoc, CollabSheet, Register } from "./types";

export interface DroppedCell {
    sheetId: string;
    col: number;
    rank: string;
    /** The text that is gone. */
    text: string;
    /** The peer that wrote the text. */
    writtenBy: string;
    /** The peer whose delete discarded it. */
    deletedBy: string;
}

export interface MergeResult {
    doc: CollabDoc;
    dropped: DroppedCell[];
}

/** The first delete, so two concurrent deletes settle the same way on both peers. */
function firstDelete(a: Stamp | null, b: Stamp | null): Stamp | null {
    if (a === null) return b;
    if (b === null) return a;
    return compareStamps(a, b) <= 0 ? a : b;
}

function mergeRegisters(
    a: Record<string, Register>,
    b: Record<string, Register>,
): Record<string, Register> {
    const out: Record<string, Register> = { ...a };
    for (const [path, reg] of Object.entries(b)) {
        const mine = out[path];
        out[path] = mine && compareStamps(mine.stamp, reg.stamp) >= 0 ? mine : reg;
    }
    return out;
}

function mergeCell(a: CollabCell, b: CollabCell): CollabCell {
    const text = compareStamps(a.textStamp, b.textStamp) >= 0 ? a : b;
    const meta = compareStamps(a.metaStamp, b.metaStamp) >= 0 ? a : b;
    return {
        col: a.col,
        rank: a.rank,
        actor: a.actor,
        text: text.text,
        textStamp: text.textStamp,
        meta: meta.meta,
        metaStamp: meta.metaStamp,
        deleted: firstDelete(a.deleted, b.deleted),
    };
}

function mergeSheet(
    sheetId: string,
    local: CollabSheet | undefined,
    incoming: CollabSheet,
    dropped: DroppedCell[],
): CollabSheet {
    if (!local) return incoming;
    const cells: Record<string, CollabCell> = { ...local.cells };
    const buried: DroppedCell[] = [];
    for (const [key, remote] of Object.entries(incoming.cells)) {
        const mine = cells[key];
        const merged = mine ? mergeCell(mine, remote) : remote;
        cells[key] = merged;
        // A cell this replica held alive, with text, that the merge just
        // buried. The peer that typed it is the one that must be told.
        if (
            mine &&
            mine.deleted === null &&
            merged.deleted !== null &&
            (merged.text ?? "").trim() !== ""
        ) {
            buried.push({
                sheetId,
                col: merged.col,
                rank: merged.rank,
                text: merged.text as string,
                writtenBy: merged.textStamp.actor,
                deletedBy: merged.deleted.actor,
            });
        }
    }
    // The report reads in grid order, not in whatever order the keys arrived.
    buried.sort((x, y) => x.col - y.col || (x.rank < y.rank ? -1 : x.rank > y.rank ? 1 : 0));
    dropped.push(...buried);
    return {
        id: sheetId,
        fields: mergeRegisters(local.fields, incoming.fields),
        deleted: firstDelete(local.deleted, incoming.deleted),
        cells,
    };
}

export function merge(local: CollabDoc, incoming: CollabDoc): MergeResult {
    const dropped: DroppedCell[] = [];
    const sheets: Record<string, CollabSheet> = { ...local.sheets };
    for (const [sheetId, remote] of Object.entries(incoming.sheets)) {
        sheets[sheetId] = mergeSheet(sheetId, sheets[sheetId], remote, dropped);
    }
    return {
        doc: {
            roundId: local.roundId,
            round: mergeRegisters(local.round, incoming.round),
            sheets,
        },
        dropped,
    };
}
