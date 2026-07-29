/**
 * Seeding a replica from a file, and projecting one back into a round.
 *
 * Seeding is a pure function of the file: row `i` of every column gets
 * `seedRank(i)`, every value gets the origin stamp, and every seeded cell is
 * credited to no actor. Two peers that open one round therefore hold
 * byte-identical replicas before they exchange a single message, which is what
 * makes the first merge correct instead of a duplication.
 */

import {
    compareSheets,
    emptyScouting,
    type CellMeta,
    type FlowRound,
    type FlowSheet,
} from "@/lib/model/flow";
import type { Scouting } from "@/lib/model/types";

import { seedRank } from "./rank";
import { ORIGIN_STAMP, type Stamp } from "./stamp";
import {
    cellKey,
    compareCells,
    flattenLeaves,
    setPath,
    type CollabCell,
    type CollabDoc,
    type CollabSheet,
    type Json,
    type Register,
} from "./types";

/** Fields of a sheet the replica stores as cells, not as registers. */
const SHEET_GRID_FIELDS: Record<string, true> = { id: true, data: true, meta: true };
/** Fields of a round that stay local: the file's own bookkeeping. */
const ROUND_LOCAL_FIELDS: Record<string, true> = {
    id: true,
    createdAt: true,
    updatedAt: true,
    sheets: true,
};

/** A sheet's live cells in one column, in row order. */
export function liveCells(sheet: CollabSheet, col: number): CollabCell[] {
    return Object.values(sheet.cells)
        .filter((c) => c.col === col && c.deleted === null)
        .sort(compareCells);
}

/**
 * The widest a sheet gets. A cell's `col` arrives inside a peer's document and
 * becomes the bound of every loop that walks a sheet's columns: the projection
 * here, the grid patch, and the local row insert and remove. One cell naming a
 * column index far out asks for more array slots than the machine has memory
 * for, and the loop never returns. A flow is a speech per column.
 */
const MAX_COL = 512;

/** One past the highest column index the sheet holds a cell in. */
export function sheetWidth(sheet: CollabSheet): number {
    let width = 0;
    for (const cell of Object.values(sheet.cells)) {
        // A cell outside the range projects nowhere, which is already what a
        // cell nobody can see looks like.
        if (!Number.isInteger(cell.col) || cell.col < 0 || cell.col >= MAX_COL) continue;
        width = Math.max(width, cell.col + 1);
    }
    return width;
}

export function seedSheet(sheet: FlowSheet, stamp: Stamp): CollabSheet {
    const leaves: Record<string, Json> = {};
    for (const [key, value] of Object.entries(sheet)) {
        if (SHEET_GRID_FIELDS[key]) continue;
        flattenLeaves(value, key, leaves);
    }
    const fields: Record<string, Register> = {};
    for (const [path, value] of Object.entries(leaves)) fields[path] = { value, stamp };

    // The grid is a rectangle even when the stored rows are ragged, so the
    // replica seeds the rectangle and a short row gains empty cells.
    const width = sheet.data.reduce((w, row) => Math.max(w, row.length), 0);
    const cells: Record<string, CollabCell> = {};
    sheet.data.forEach((row, rowIndex) => {
        const rank = seedRank(rowIndex);
        for (let col = 0; col < width; col++) {
            const meta = sheet.meta[`${rowIndex},${col}`];
            cells[cellKey(col, rank, "")] = {
                col,
                rank,
                actor: "",
                text: row[col] ?? null,
                textStamp: stamp,
                meta: meta ? ({ ...meta } as Record<string, Json>) : {},
                metaStamp: stamp,
                deleted: null,
            };
        }
    });
    return { id: sheet.id, fields, deleted: null, cells };
}

export function seedDoc(round: FlowRound): CollabDoc {
    const leaves: Record<string, Json> = {};
    for (const [key, value] of Object.entries(round)) {
        if (ROUND_LOCAL_FIELDS[key]) continue;
        flattenLeaves(value, key, leaves);
    }
    const roundRegisters: Record<string, Register> = {};
    for (const [path, value] of Object.entries(leaves)) {
        roundRegisters[path] = { value, stamp: ORIGIN_STAMP };
    }
    const sheets: Record<string, CollabSheet> = {};
    for (const sheet of round.sheets) sheets[sheet.id] = seedSheet(sheet, ORIGIN_STAMP);
    return { roundId: round.id, round: roundRegisters, sheets };
}

export function projectSheet(sheet: CollabSheet): FlowSheet {
    const shape: Record<string, unknown> = {};
    for (const [path, reg] of Object.entries(sheet.fields)) setPath(shape, path, reg.value);

    const width = sheetWidth(sheet);
    const columns: CollabCell[][] = [];
    let height = 0;
    for (let col = 0; col < width; col++) {
        const live = liveCells(sheet, col);
        columns.push(live);
        height = Math.max(height, live.length);
    }
    const data: (string | null)[][] = [];
    const meta: Record<string, CellMeta> = {};
    for (let row = 0; row < height; row++) {
        const line: (string | null)[] = [];
        for (let col = 0; col < width; col++) {
            const cell = columns[col][row];
            line.push(cell?.text ?? null);
            if (cell && Object.keys(cell.meta).length > 0) {
                meta[`${row},${col}`] = cell.meta as CellMeta;
            }
        }
        data.push(line);
    }
    return { ...(shape as Omit<FlowSheet, "id" | "data" | "meta">), id: sheet.id, data, meta };
}

/**
 * The round a replica describes. `base` supplies the two fields the replica
 * deliberately does not carry, so a partner's clock never rewrites this file's
 * creation time.
 *
 * `settled` is the document `base` was last projected from, when the caller
 * has one. A merge builds a new object only for the sheets it touched, so a
 * sheet whose replica is the very same object describes the very same sheet,
 * and the copy already in `base` stands rather than being derived again.
 * Without that, one cell arriving from a partner re-derives every sheet in the
 * round, up to thirty times a second while they type.
 */
export function projectDoc(doc: CollabDoc, base: FlowRound, settled?: CollabDoc): FlowRound {
    const shape: Record<string, unknown> = {};
    for (const [path, reg] of Object.entries(doc.round)) setPath(shape, path, reg.value);
    const already = new Map(base.sheets.map((s) => [s.id, s]));
    const sheets = Object.values(doc.sheets)
        .filter((s) => s.deleted === null)
        .map((sheet) => {
            const untouched =
                settled?.sheets[sheet.id] === sheet ? already.get(sheet.id) : undefined;
            return untouched ?? projectSheet(sheet);
        })
        .sort(compareSheets);
    return {
        ...(shape as Omit<FlowRound, "id" | "createdAt" | "updatedAt" | "scouting" | "sheets">),
        id: doc.roundId,
        createdAt: base.createdAt,
        updatedAt: base.updatedAt,
        scouting: (shape.scouting as Scouting) ?? emptyScouting(),
        sheets,
    };
}
