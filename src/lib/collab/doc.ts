/**
 * Seeding a replica from a file, and projecting one back into a round.
 *
 * Seeding is a pure function of the file: row `i` of every column gets
 * `seedRank(i)`, every value gets the origin stamp, and every seeded cell is
 * credited to no actor. Two peers that open one round therefore hold
 * byte-identical replicas before they exchange a single message, which is what
 * makes the first merge correct instead of a duplication.
 */

import { getEvent } from "@/lib/format/events";
import {
    compareSheets,
    emptyScouting,
    type CellMeta,
    type FlowRound,
    type FlowSheet,
} from "@/lib/model/flow";
import type { Scouting } from "@/lib/model/types";
import {
    MAX_ROUND_CELLS,
    holdsCellMeta,
    holdsScouting,
    paddedCells,
} from "@/lib/persistence/flowFile";

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

/**
 * The widest a sheet gets, and the tallest a column gets. Both coordinates
 * arrive inside a peer's document and both become the bound of a loop that
 * walks a sheet: the projection here, the grid patch, and the local row insert
 * and remove. The projection materializes their product, so bounding one alone
 * buys nothing - four thousand cells in one column asks for as many array slots
 * as a column index far out does. A flow is a speech per column, a speech is
 * not two thousand lines, and one sheet's product stays under the
 * MAX_ROUND_CELLS the file enforces on read. Two sheets' products do not, which
 * is why the round's own budget is spent in `projectDoc` rather than here.
 */
const MAX_COL = 512;
const MAX_ROWS = 2048;

/** A sheet's live cells in one column, in row order. */
export function liveCells(sheet: CollabSheet, col: number): CollabCell[] {
    const live = Object.values(sheet.cells)
        .filter((c) => c.col === col && c.deleted === null)
        .sort(compareCells);
    // A row past the bound projects nowhere, which is already what a cell
    // nobody can see looks like.
    return live.length > MAX_ROWS ? live.slice(0, MAX_ROWS) : live;
}

/** One past the highest column index the sheet holds a cell in. */
export function sheetWidth(sheet: CollabSheet): number {
    let width = 0;
    for (const cell of Object.values(sheet.cells)) {
        // A cell outside the range projects nowhere, for the same reason.
        if (!Number.isInteger(cell.col) || cell.col < 0 || cell.col >= MAX_COL) continue;
        width = Math.max(width, cell.col + 1);
    }
    return width;
}

/**
 * The cells a sheet's grid pads out to, which is what the file counts it as.
 * The same number `projectSheet` goes on to produce, derived without
 * materializing the rectangle, because `projectDoc` has to know what a sheet
 * costs before it can decide how much of the round's budget the sheet may have.
 */
function projectedCells(sheet: CollabSheet): number {
    const width = sheetWidth(sheet);
    let height = 0;
    for (let col = 0; col < width; col++) height = Math.max(height, liveCells(sheet, col).length);
    return width * height;
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

/**
 * The sheet a replica describes.
 *
 * Every field but the grid is a replicated register holding whatever a peer put
 * on the wire, and each one below is written to the file, whose parser refuses a
 * title that is not text, a group that is not a side, or an order that is not a
 * number. A round the parser refuses is a round the debater loses, so a value
 * the file cannot hold falls back to what a file predating that field gets.
 *
 * `room` is the cells of the round's budget this sheet may claim; see
 * `projectDoc`. One sheet on its own cannot reach the default, because MAX_COL
 * times MAX_ROWS is under it.
 */
export function projectSheet(sheet: CollabSheet, room = MAX_ROUND_CELLS): FlowSheet {
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
    // The file counts a sheet as its rows times its widest row, so the rows
    // that fit in `room` are room / width. Rows come off the bottom, which is
    // the end of the flow a sheet the format cannot hold has run past, and the
    // least of that sheet to lose: dropping columns would lose whole speeches
    // and dropping the sheet would lose one the debater may be looking at.
    if (width > 0) height = Math.min(height, Math.floor(room / width));
    const data: (string | null)[][] = [];
    const meta: Record<string, CellMeta> = {};
    for (let row = 0; row < height; row++) {
        const line: (string | null)[] = [];
        for (let col = 0; col < width; col++) {
            const cell = columns[col][row];
            line.push(typeof cell?.text === "string" ? cell.text : null);
            if (cell && Object.keys(cell.meta).length > 0 && holdsCellMeta(cell.meta)) {
                meta[`${row},${col}`] = cell.meta as CellMeta;
            }
        }
        data.push(line);
    }
    const projected: FlowSheet = {
        ...(shape as Omit<FlowSheet, "id" | "data" | "meta">),
        id: sheet.id,
        title: typeof shape.title === "string" ? shape.title : "",
        group: shape.group === "neg" ? "neg" : "aff",
        order: typeof shape.order === "number" && Number.isFinite(shape.order) ? shape.order : 0,
        kind: shape.kind === "cx" ? "cx" : "flow",
        data,
        meta,
    };
    if (typeof projected.startSpeechId !== "string") delete projected.startSpeechId;
    return projected;
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
 *
 * This is also where the file's cell ceiling is spent, because the ceiling is a
 * total across sheets while the merge's `MAX_CELLS` is per sheet: two sheets a
 * peer grew to the widest and tallest this build projects sum past it, and then
 * every autosave of that round is refused for as long as the round holds them.
 * The budget belongs here and not in the merge - what the merge accepts must
 * not depend on the replica's own other sheets, or two peers holding different
 * sheets would accept different cells and diverge. A projection is local: the
 * replica still holds every cell it was sent, and only the file is bounded.
 *
 * Cheapest sheet first, each offered everything still unspent, so a round the
 * format holds is projected exactly as it would be with no budget at all: the
 * budget only ever runs out on a sheet larger than every sheet before it. A
 * sheet is therefore clamped only once the sheets no larger than it have spent
 * the whole two million, which at the 512-sheet ceiling takes 3,907 cells each
 * - a 325-row sheet of twelve speeches. A fat real round is a few hundred rows
 * by a dozen speeches per sheet, three orders of magnitude below that.
 */
export function projectDoc(doc: CollabDoc, base: FlowRound, settled?: CollabDoc): FlowRound {
    const shape: Record<string, unknown> = {};
    for (const [path, reg] of Object.entries(doc.round)) setPath(shape, path, reg.value);
    const already = new Map(base.sheets.map((s) => [s.id, s]));
    // What each sheet costs the file: the reused ones from the copy already
    // projected, the rest derived. Cheapest first, then by id, so two peers
    // holding the same document write the same file - `order` would not do,
    // being a register a peer writes.
    const costed = Object.values(doc.sheets)
        .filter((s) => s.deleted === null)
        .map((sheet) => {
            const untouched =
                settled?.sheets[sheet.id] === sheet ? already.get(sheet.id) : undefined;
            const cost = untouched ? paddedCells(untouched.data) : projectedCells(sheet);
            return { sheet, untouched, cost };
        })
        .sort((a, b) => a.cost - b.cost || (a.sheet.id < b.sheet.id ? -1 : 1));

    let unspent = MAX_ROUND_CELLS;
    const sheets = costed
        .map(({ sheet, untouched, cost }) => {
            if (untouched && cost <= unspent) {
                unspent -= cost;
                return untouched;
            }
            // Clamped to what is left, and counted as what it came back as.
            const projected = projectSheet(sheet, unspent);
            unspent -= paddedCells(projected.data);
            return projected;
        })
        .sort(compareSheets);
    return {
        ...(shape as Omit<FlowRound, "id" | "createdAt" | "updatedAt" | "scouting" | "sheets">),
        id: doc.roundId,
        createdAt: base.createdAt,
        updatedAt: base.updatedAt,
        // Replicated registers, so a peer chooses all three, and all three are
        // written to the file. `getEvent` already names the fallback an
        // unknown event gets on read; this is the same fallback on write.
        event: getEvent(shape.event as string).id,
        firstSide: shape.firstSide === "neg" ? "neg" : "aff",
        scouting: holdsScouting(shape.scouting) ? (shape.scouting as Scouting) : emptyScouting(),
        sheets,
    };
}
