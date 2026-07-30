/**
 * The slices of Handsontable the grid tests stand in for.
 *
 * A real Handsontable needs a DOM container and a full settings object, so
 * these fakes cover exactly the surface the module under test touches, backed
 * by plain arrays and maps a test can assert on directly.
 */

import { vi } from "vitest";

import type { CellChange } from "@/lib/grid/cellShift";
import type { MoveGrid } from "@/lib/grid/moveSession";
import type { CellSource } from "@/lib/model/flow";

/** The two meta keys a flow cell carries. */
export interface FakeCellMeta {
    className?: string;
    source?: CellSource;
}

export interface MetaStore {
    /** The cell's meta, created empty on first touch, live for reading back. */
    at(row: number, col: number): FakeCellMeta;
    getCellMeta(row: number, col: number): FakeCellMeta;
    setCellMeta(row: number, col: number, key: string, value: unknown): void;
}

/** A className/source store keyed "row,col", optionally seeded per cell. */
export function metaStore(seed: Iterable<readonly [string, FakeCellMeta]> = []): MetaStore {
    const meta = new Map<string, FakeCellMeta>(seed);
    const at = (row: number, col: number) => {
        const key = `${row},${col}`;
        if (!meta.has(key)) meta.set(key, {});
        return meta.get(key)!;
    };
    return {
        at,
        getCellMeta: at,
        setCellMeta: (row, col, key, value) => {
            if (key === "source") at(row, col).source = value as CellSource | undefined;
            else at(row, col).className = value as string;
        },
    };
}

/**
 * A grid whose selection spans rows 0..lastRow of column 0 - the shape the
 * decoration commands read, and all they read.
 */
export function selectionHot(lastRow: number, meta: MetaStore = metaStore()) {
    return {
        getSelectedRange: () => [
            {
                highlight: { row: 0, col: 0 },
                getTopLeftCorner: () => ({ row: 0, col: 0 }),
                getBottomRightCorner: () => ({ row: lastRow, col: 0 }),
            },
        ],
        getCellMeta: meta.getCellMeta,
        setCellMeta: meta.setCellMeta,
        render: vi.fn(),
    };
}

/**
 * A column-major grid over a plain array: what cellShift and moveSession read
 * and write. `col` picks one column's text out as an array, which is how the
 * rotation assertions read.
 */
export function fakeGrid(
    data: (string | null)[][],
    classNames: Record<string, string> = {},
    sources: Record<string, CellSource> = {},
): MoveGrid & {
    data: (string | null)[][];
    classNames: Record<string, string>;
    sources: Record<string, CellSource>;
    col(c: number): (string | null)[];
} {
    const store = { ...classNames };
    const srcStore = { ...sources };
    return {
        data,
        classNames: store,
        sources: srcStore,
        countRows: () => data.length,
        countCols: () => data[0]?.length ?? 0,
        getDataAtCell: (r, c) => data[r][c],
        setDataAtCell: (changes: CellChange[]) => {
            for (const [r, c, v] of changes) data[r][c] = v;
        },
        getCellMeta: (r, c) => ({ className: store[`${r},${c}`], source: srcStore[`${r},${c}`] }),
        setCellMeta: (r, c, key, value) => {
            const cell = `${r},${c}`;
            if (key === "source") {
                if (value) srcStore[cell] = value as CellSource;
                else delete srcStore[cell];
            } else if (value) {
                store[cell] = value as string;
            } else {
                delete store[cell];
            }
        },
        col(c) {
            return data.map((row) => row[c]);
        },
    };
}
