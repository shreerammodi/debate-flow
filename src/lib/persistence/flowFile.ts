/**
 * The .ebb file format: a version envelope wrapping one FlowRound, written as
 * pretty-printed JSON so the file stays diffable and readable outside ebb.
 *
 * Version 3 is the Handsontable-native model. Versions 1-2 are the legacy node
 * model and are rejected outright; they were never migratable. A file on disk
 * outlives the build that wrote it, so an older version is normalized rather
 * than refused, and only a newer one is refused - this build cannot know what
 * it would silently drop.
 *
 * Validation is strict on purpose. A database row was written by code that had
 * already type-checked it; a file can be truncated by a full disk, mangled by a
 * sync client, or hand-edited. Failing at this boundary with the path to the bad
 * value beats rendering half a round.
 */

import { EVENTS } from "@/lib/format/events";
import { normalizeFlow, type FlowRound } from "@/lib/model/flow";
import { uid } from "@/lib/model/ids";

export const FLOW_FILE_VERSION = 3;

/** Serialize a round as .ebb file text. */
export function serializeFlow(round: FlowRound): string {
    return JSON.stringify({ version: FLOW_FILE_VERSION, round }, null, 2) + "\n";
}

// --- Validation --------------------------------------------------------------

type Obj = Record<string, unknown>;

function fail(path: string, expected: string): never {
    throw new Error(`Invalid flow file: ${path} ${expected}`);
}

function obj(value: unknown, path: string): Obj {
    if (typeof value !== "object" || value === null || Array.isArray(value)) {
        fail(path, "is not an object");
    }
    return value as Obj;
}

function str(value: unknown, path: string): string {
    if (typeof value !== "string") fail(path, "is not a string");
    return value;
}

/** Absent and null both mean "unset"; anything else must be the right type. */
function optional(value: unknown): boolean {
    return value === undefined || value === null;
}

function optStr(value: unknown, path: string): void {
    if (!optional(value) && typeof value !== "string") fail(path, "is not a string");
}

function optBool(value: unknown, path: string): void {
    if (!optional(value) && typeof value !== "boolean") fail(path, "is not a boolean");
}

function finiteNum(value: unknown, path: string): number {
    if (typeof value !== "number" || !Number.isFinite(value)) fail(path, "is not a number");
    return value;
}

function checkDebater(value: unknown, path: string): void {
    const d = obj(value, path);
    str(d.first, `${path}.first`);
    str(d.last, `${path}.last`);
}

function checkScouting(value: unknown, path: string): void {
    const sc = obj(value, path);
    for (const side of ["aff", "neg"] as const) {
        const team = obj(sc[side], `${path}.${side}`);
        checkDebater(team.first, `${path}.${side}.first`);
        checkDebater(team.second, `${path}.${side}.second`);
    }
    for (const key of [
        "affSchool",
        "negSchool",
        "tournament",
        "round",
        "flight",
        "date",
        "judge",
    ]) {
        optStr(sc[key], `${path}.${key}`);
    }
    if (!optional(sc.decision)) {
        const d = obj(sc.decision, `${path}.decision`);
        if (!optional(d.vote) && d.vote !== "aff" && d.vote !== "neg") {
            fail(`${path}.decision.vote`, 'is not "aff" or "neg"');
        }
        optStr(d.rfd, `${path}.decision.rfd`);
        if (!optional(d.peerNotes)) {
            // One entry per peer, each that peer's own notes. A hand edit that
            // put something else in here would reach the RFD preview.
            const notes = obj(d.peerNotes, `${path}.decision.peerNotes`);
            for (const [endpointId, note] of Object.entries(notes)) {
                optStr(note, `${path}.decision.peerNotes.${endpointId}`);
            }
        }
    }
}

function checkCellMeta(value: unknown, path: string): void {
    const m = obj(value, path);
    optBool(m.bold, `${path}.bold`);
    optBool(m.highlight, `${path}.highlight`);
    optBool(m.card, `${path}.card`);
    optBool(m.group, `${path}.group`);
    if (!optional(m.answers)) {
        const a = obj(m.answers, `${path}.answers`);
        str(a.sheetId, `${path}.answers.sheetId`);
        finiteNum(a.row, `${path}.answers.row`);
        finiteNum(a.col, `${path}.answers.col`);
    }
    if (!optional(m.source)) {
        const s = obj(m.source, `${path}.source`);
        str(s.app, `${path}.source.app`);
        str(s.token, `${path}.source.token`);
        str(s.key, `${path}.source.key`);
        optStr(s.title, `${path}.source.title`);
    }
}

function checkSheet(value: unknown, path: string): void {
    const s = obj(value, path);
    str(s.id, `${path}.id`);
    str(s.title, `${path}.title`);
    if (s.group !== "aff" && s.group !== "neg") fail(`${path}.group`, 'is not "aff" or "neg"');
    finiteNum(s.order, `${path}.order`);
    if (!optional(s.kind) && s.kind !== "flow" && s.kind !== "cx") {
        fail(`${path}.kind`, 'is not "flow" or "cx"');
    }
    optStr(s.startSpeechId, `${path}.startSpeechId`);

    if (!Array.isArray(s.data)) fail(`${path}.data`, "is not an array");
    s.data.forEach((row, r) => {
        if (!Array.isArray(row)) fail(`${path}.data[${r}]`, "is not a row");
        row.forEach((cell, c) => {
            if (cell !== null && typeof cell !== "string") {
                fail(`${path}.data[${r}][${c}]`, "is not text or null");
            }
        });
    });

    // Sparse and optional: an older sheet may predate cell metadata entirely.
    if (!optional(s.meta)) {
        const meta = obj(s.meta, `${path}.meta`);
        for (const key of Object.keys(meta)) checkCellMeta(meta[key], `${path}.meta["${key}"]`);
    }
}

/** Validate a parsed round, throwing with the path to the first bad value. */
function checkRound(value: unknown, path: string): FlowRound {
    const r = obj(value, path);
    str(r.id, `${path}.id`);
    finiteNum(r.createdAt, `${path}.createdAt`);
    finiteNum(r.updatedAt, `${path}.updatedAt`);
    if (!optional(r.event)) {
        const event = str(r.event, `${path}.event`);
        if (!(event in EVENTS)) fail(`${path}.event`, "is not a known debate event");
    }
    if (!optional(r.firstSide) && r.firstSide !== "aff" && r.firstSide !== "neg") {
        fail(`${path}.firstSide`, 'is not "aff" or "neg"');
    }
    checkScouting(r.scouting, `${path}.scouting`);
    if (!Array.isArray(r.sheets)) fail(`${path}.sheets`, "is not an array");
    r.sheets.forEach((s, i) => checkSheet(s, `${path}.sheets[${i}]`));
    return value as FlowRound;
}

// --- Reading -----------------------------------------------------------------

function parseEnvelope(text: string): Obj {
    let parsed: unknown;
    try {
        parsed = JSON.parse(text);
    } catch {
        throw new Error("Not a flow file: the contents are not valid JSON");
    }
    const envelope = obj(parsed, "the file");
    const version = finiteNum(envelope.version, "the file version");
    if (version > FLOW_FILE_VERSION) {
        throw new Error(
            `This flow was written by a newer version of ebb (file version ${version}). Update ebb to open it.`,
        );
    }
    if (version < FLOW_FILE_VERSION) {
        throw new Error(
            `Flow file version ${version} is from a retired format and cannot be opened.`,
        );
    }
    return envelope;
}

/**
 * Parse .ebb file text into the round it holds, preserving its identity.
 * Opening a file is not importing one: the path is the identity now, so the
 * round's own id, createdAt, and history survive the round trip.
 */
export function parseFlowFile(text: string): FlowRound {
    const envelope = parseEnvelope(text);
    if (envelope.kind === "backup") {
        throw new Error("That is a multi-flow backup, not a single flow.");
    }
    return normalizeFlow(checkRound(envelope.round, "round"));
}

/**
 * Parse a legacy export - either a single `{version, round}` or a
 * `{version, kind:"backup", rounds}` - into rounds with fresh identities.
 * These files were snapshots rather than documents, so materializing one into
 * the flows folder mints a new identity per round the way importing always did.
 */
export function parseLegacyExport(text: string): FlowRound[] {
    const envelope = parseEnvelope(text);
    const backup = envelope.kind === "backup";
    if (backup && !Array.isArray(envelope.rounds)) fail("rounds", "is not an array");
    const raw = backup ? (envelope.rounds as unknown[]) : [envelope.round];

    const now = Date.now();
    return raw.map((r, i) => ({
        ...normalizeFlow(checkRound(r, backup ? `rounds[${i}]` : "round")),
        id: uid("round"),
        createdAt: now,
        updatedAt: now,
    }));
}
