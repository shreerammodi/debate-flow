/**
 * Handsontable-native round model. Each sheet stores its grid as a 2D array of
 * cell text plus sparse per-cell metadata; columns are never stored (they
 * derive from the round's event definition and the sheet's startSpeechId slice).
 */

import { getEvent, type EventId } from "@/lib/format/events";
import { uid } from "@/lib/model/ids";
import type { Scouting, Side } from "@/lib/model/types";

/** Where this cell's text came from, when it was sent in from another app. */
export interface CellSource {
    /** Origin app id from the cardmirror-bridge handshake, e.g. "cardmirror". */
    app: string;
    /** The origin's opaque provenance token, handed back verbatim to jump. */
    token: string;
    /** Stable equality key the origin mints, used for forward search. */
    key: string;
    /** Origin document title, for messages like "open X first". */
    title?: string;
}

export interface CellMeta {
    bold?: boolean;
    highlight?: boolean;
    /** Tags the cell as a card (a piece of evidence). */
    card?: boolean;
    /** Marks the cell as part of a visual group (a left bar hugging the run). */
    group?: boolean;
    /** Marks the argument dead: kicked, turned away, or never extended. */
    kicked?: boolean;
    /** Reserved for the links phase; nothing reads or writes it yet. */
    answers?: { sheetId: string; row: number; col: number };
    /** Provenance for text handed in by another app; absent for typed cells. */
    source?: CellSource;
}

export interface FlowSheet {
    id: string;
    title: string;
    group: "aff" | "neg";
    order: number;
    /** Absent / "flow" = argument grid. "cx" = the cross-ex sheet. */
    kind?: "flow" | "cx";
    /** Leftmost speech column shown (absent = the side's first speech in the round's event). */
    startSpeechId?: string;
    /** rows x speech-columns cell text, Handsontable-native. */
    data: (string | null)[][];
    /** Sparse per-cell metadata keyed "row,col". */
    meta: Record<string, CellMeta>;
}

export interface FlowRound {
    id: string;
    createdAt: number;
    updatedAt: number;
    /** Debate event; absent (legacy rounds) = "policy". */
    event?: EventId;
    /** First-speaking side; meaningful only for variable-order events (PF). */
    firstSide?: Side;
    scouting: Scouting;
    sheets: FlowSheet[];
}

const emptyDebater = () => ({ first: "", last: "" });

export function emptyScouting(): Scouting {
    return {
        aff: { first: emptyDebater(), second: emptyDebater() },
        neg: { first: emptyDebater(), second: emptyDebater() },
    };
}

export function makeFlowSheet(input: {
    title: string;
    group: "aff" | "neg";
    order: number;
}): FlowSheet {
    return {
        id: uid("sheet"),
        title: input.title,
        group: input.group,
        order: input.order,
        kind: "flow",
        data: [],
        meta: {},
    };
}

/** The pinned cross-examination sheet. order = -1 so it sorts above flow sheets. */
export function makeCxFlowSheet(title = "CX"): FlowSheet {
    return {
        id: uid("sheet"),
        title,
        group: "aff",
        order: -1,
        kind: "cx",
        data: [],
        meta: {},
    };
}

export function makeFlowRound(input: { event?: EventId; firstSide?: Side } = {}): FlowRound {
    const now = Date.now();
    const event = input.event ?? "policy";
    const firstSide = input.firstSide ?? "aff";
    const crossEx = getEvent(event).crossEx;
    return {
        id: uid("round"),
        createdAt: now,
        updatedAt: now,
        event,
        firstSide,
        scouting: emptyScouting(),
        sheets: [
            // Parliamentary has no cross-examination, so it opens with no
            // cross-ex sheet to leave empty.
            ...(crossEx ? [makeCxFlowSheet(crossEx.title)] : []),
            // The first sheet belongs to whoever speaks first, so the round
            // opens on the constructive that actually starts it.
            makeFlowSheet({ title: "1.", group: firstSide, order: 0 }),
        ],
    };
}

/**
 * Fill defaults on a round read from a file. Never mutates input. Drops the
 * legacy `deletedAt` field, which soft-deleted a round back when flows lived in
 * a database; a flow is now a file, and the filesystem owns deletion.
 */
export function normalizeFlow(raw: FlowRound): FlowRound {
    const { deletedAt: _legacyDeletedAt, ...rest } = raw as FlowRound & { deletedAt?: unknown };
    const r: FlowRound = {
        ...rest,
        event: raw.event ?? "policy",
        firstSide: raw.firstSide ?? "aff",
        scouting: raw.scouting ? { ...raw.scouting } : emptyScouting(),
        sheets: (raw.sheets ?? []).map((s) => ({
            ...s,
            kind: s.kind ?? "flow",
            data: Array.isArray(s.data) ? s.data : [],
            meta: s.meta ?? {},
        })),
    };
    const crossEx = getEvent(r.event).crossEx;
    if (crossEx && !r.sheets.some((s) => s.kind === "cx")) {
        r.sheets = [makeCxFlowSheet(crossEx.title), ...r.sheets];
    }
    return r;
}

/**
 * Total order on sheets. `reorderSheets` renumbers to contiguous integers, so
 * two peers reordering at once can produce one order value twice; resolving
 * that tie by array position would differ per peer and the two sidebars would
 * visibly disagree.
 */
export function compareSheets(a: FlowSheet, b: FlowSheet): number {
    if (a.order !== b.order) return a.order - b.order;
    return a.id < b.id ? -1 : a.id > b.id ? 1 : 0;
}

/** Sheets sorted ascending by order (CX first at order -1). */
export function sortedSheets(round: FlowRound): FlowSheet[] {
    return round.sheets.slice().sort(compareSheets);
}

/** First flow (non-CX) sheet id by order, else the first sheet, else null. */
export function firstFlowSheetId(round: FlowRound): string | null {
    const sheets = sortedSheets(round);
    return (sheets.find((s) => s.kind !== "cx") ?? sheets[0])?.id ?? null;
}
