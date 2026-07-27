/**
 * Rewriting a run of one column as ops a peer can apply.
 *
 * A block move reorders cells inside a column, which the op union has no
 * single member for. Re-deriving the sheet instead would re-key every cell
 * from its row position, and a peer holding the old keys would not agree with
 * any of it. So the moved run is expressed the long way: take the old cells
 * out, put fresh ones in, and write the text. Every step is an op that
 * travels, which is the whole difference.
 */

import type { CollabOp } from "./ops";

/**
 * Ops that leave rows `[at, at + texts.length)` of `col` holding `texts`.
 *
 * The removes all name the same index because each one closes the gap behind
 * it, and the inserts all name the same index because each one pushes the
 * previous down.
 */
export function replaceSpanOps(
    sheetId: string,
    col: number,
    at: number,
    texts: readonly (string | null)[],
): CollabOp[] {
    if (texts.length === 0) return [];

    const ops: CollabOp[] = [];
    for (let i = 0; i < texts.length; i++) {
        ops.push({ kind: "removeCell", sheetId, col, row: at });
    }
    for (let i = 0; i < texts.length; i++) {
        ops.push({ kind: "insertCell", sheetId, col, row: at });
    }
    texts.forEach((text, i) => {
        ops.push({ kind: "cellText", sheetId, col, row: at + i, text });
    });
    return ops;
}
