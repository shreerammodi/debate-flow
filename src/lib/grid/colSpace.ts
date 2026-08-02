/**
 * The two column spaces a padded pane holds at once.
 *
 * A model column indexes a sheet's stored data, and is what the file holds,
 * what an op carries on the wire, what a peer's cursor names, what cell search
 * returns and what the exporter reads. A grid column indexes the Handsontable
 * instance, which on an aligned pane leads with one inert column per speech
 * the sheet does not show.
 *
 * They are branded so the two cannot be mixed without a conversion the
 * compiler demands: a boundary added later is a type error rather than
 * something a reviewer has to notice. The brands erase at runtime, so nothing
 * on the wire or on disk changes shape.
 */

export type ModelCol = number & { readonly __modelCol: unique symbol };
export type GridCol = number & { readonly __gridCol: unique symbol };

/**
 * Names a bare number as a model column. The seam for values that arrive
 * already validated: off the wire, out of a file, or out of a store field.
 */
export function modelCol(n: number): ModelCol {
    return n as ModelCol;
}

/**
 * Names a bare number as a grid column. The seam for values Handsontable
 * hands back, whose API is untyped; this is the only place a bare number
 * becomes a grid column, so the casts are greppable.
 */
export function gridCol(n: number): GridCol {
    return n as GridCol;
}

/**
 * The cell a grid column points at. Clamped at zero: a column inside the pad
 * belongs to no cell of the sheet, and a negative index would read a row from
 * its end instead of failing.
 */
export function toModelCol(col: GridCol, spacers: number): ModelCol {
    return Math.max(col - spacers, 0) as ModelCol;
}

/** Where a cell sits on a pane carrying `spacers` inert columns. */
export function toGridCol(col: ModelCol, spacers: number): GridCol {
    return (col + spacers) as GridCol;
}
