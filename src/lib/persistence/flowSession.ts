/**
 * Opening, creating, and saving the one flow file ebb has open.
 *
 * Autosave keeps the shape it had against the database - a 500ms debounce, a
 * sequence guard so a slow earlier write cannot report over a newer one, and a
 * flush on teardown - because that behavior is what makes losing a round
 * impossible, and none of it depended on the storage being a database. Only the
 * sink changed, and it gained an atomic write on the way.
 *
 * A new flow is written to the flows folder the moment it is created, so there
 * is never an unanchored buffer, never a dirty state, and never a save prompt
 * between the user and a speech that has already started.
 */

import type { StoreApi } from "zustand";

import type { FlowRound } from "@/lib/model/flow";

import { parseFlowFile, parseLegacyExport, serializeFlow } from "./flowFile";
import { getFlowFs, type FlowFs } from "./flowFs";
import { EBB_EXT, basename, suggestFilename } from "./flowPaths";
import { dropRecent, loadRecents, promoteRecent, saveRecents } from "./recents";

/** Lifecycle of a single save, reported so the header can reassure the user. */
export type SaveStatus = "saving" | "saved" | "error";

const DEBOUNCE_MS = 500;

// --- Recents bookkeeping -------------------------------------------------------

/** Record a path as the most recently opened flow. */
export async function noteOpened(path: string, fs?: FlowFs): Promise<void> {
    const io = fs ?? (await getFlowFs());
    await saveRecents(io, promoteRecent(await loadRecents(io), path, Date.now()));
}

/** Forget a path without touching the file it points at. */
export async function forgetRecent(path: string, fs?: FlowFs): Promise<void> {
    const io = fs ?? (await getFlowFs());
    await saveRecents(io, dropRecent(await loadRecents(io), path));
}

// --- Reading -------------------------------------------------------------------

/**
 * Read the flow at `path`, or null when the file is gone - an ordinary outcome
 * for a recent entry whose flow was moved or deleted outside ebb. A file that
 * exists but does not parse throws, because that is a real problem the user
 * needs told about rather than a flow quietly vanishing from the list.
 */
export async function readFlowAt(path: string, fs?: FlowFs): Promise<FlowRound | null> {
    const io = fs ?? (await getFlowFs());
    const text = await io.readFlow(path);
    return text === null ? null : parseFlowFile(text);
}

// --- Creating and saving -------------------------------------------------------

/** Write a brand-new flow into the flows folder; resolves to the path used. */
export async function createFlowFile(round: FlowRound, fs?: FlowFs): Promise<string> {
    const io = fs ?? (await getFlowFs());
    const { flowsDir } = await io.locations();
    const path = await io.createFlow(flowsDir, suggestFilename(round), serializeFlow(round));
    await noteOpened(path, io);
    return path;
}

/**
 * Pick a flow to open. A `.ebb` is opened in place; anything else is treated as
 * a legacy export and materialized into the flows folder first, so the exports
 * users already have stay openable without becoming a second kind of document.
 * Resolves to the path to route to, or null when the picker was cancelled.
 */
export async function pickFlowToOpen(fs?: FlowFs): Promise<string | null> {
    const io = fs ?? (await getFlowFs());
    const picked = await io.pickOpenPath();
    if (!picked) return null;
    if (picked.toLowerCase().endsWith(EBB_EXT)) {
        await noteOpened(picked, io);
        return picked;
    }
    return importLegacyExport(picked, io);
}

/**
 * Turn a pre-.ebb JSON export into real flow files. A backup holding many
 * rounds becomes many files; the first is the one to open.
 */
async function importLegacyExport(path: string, io: FlowFs): Promise<string> {
    const text = await io.readFlow(path);
    if (text === null) throw new Error(`${basename(path)} no longer exists`);

    const rounds = parseLegacyExport(text);
    if (!rounds.length) throw new Error(`${basename(path)} holds no flows`);

    const { flowsDir } = await io.locations();
    const written: string[] = [];
    for (const round of rounds) {
        written.push(await io.createFlow(flowsDir, suggestFilename(round), serializeFlow(round)));
    }
    await noteOpened(written[0], io);
    return written[0];
}

/**
 * Save the open round to a location the user picks, and continue editing there.
 * Resolves to the new path, or null when the picker was cancelled.
 */
export async function saveFlowAs(round: FlowRound, fs?: FlowFs): Promise<string | null> {
    const io = fs ?? (await getFlowFs());
    const path = await io.pickSavePath(suggestFilename(round));
    if (!path) return null;
    await io.writeFlow(path, serializeFlow(round));
    await noteOpened(path, io);
    return path;
}

/** Write immediately, reporting the outcome. Backs the manual retry affordance. */
export async function saveFlowNow(
    path: string,
    round: FlowRound,
    onStatus?: (status: SaveStatus) => void,
): Promise<void> {
    onStatus?.("saving");
    try {
        const io = await getFlowFs();
        await io.writeFlow(path, serializeFlow(round));
        onStatus?.("saved");
    } catch {
        onStatus?.("error");
    }
}

// --- Autosave --------------------------------------------------------------------

/**
 * Subscribe to a store holding the open round and its path, and write on every
 * change, debounced. Only the newest write reports a terminal status, so a slow
 * earlier one cannot clobber a newer one's result. The returned unsubscribe
 * flushes anything pending, so navigating away never drops the last edit.
 */
export function attachFlowAutosave(
    store: StoreApi<{ round: FlowRound | null; docPath: string | null }>,
    onStatus?: (status: SaveStatus) => void,
): () => void {
    let timer: ReturnType<typeof setTimeout> | undefined;
    let lastSeenId: string | null = null;
    let lastSeenUpdatedAt: number | null = null;
    let pending: { path: string; round: FlowRound } | null = null;
    let saveSeq = 0;

    function doSave(job: { path: string; round: FlowRound }) {
        const seq = ++saveSeq;
        onStatus?.("saving");
        getFlowFs()
            .then((io) => io.writeFlow(job.path, serializeFlow(job.round)))
            .then(
                () => {
                    if (seq === saveSeq) onStatus?.("saved");
                },
                () => {
                    if (seq === saveSeq) onStatus?.("error");
                },
            );
    }

    function flush() {
        clearTimeout(timer);
        timer = undefined;
        if (pending !== null) {
            const job = pending;
            pending = null;
            doSave(job);
        }
    }

    const unsubscribe = store.subscribe((state) => {
        const { round, docPath } = state;
        if (!round || !docPath) return;
        if (round.id === lastSeenId && round.updatedAt === lastSeenUpdatedAt) return;

        // Every round in the store arrived from the file it would be written
        // back to, so the notification that introduces one is not an edit.
        // Saving it would rewrite identical bytes on every open and touch the
        // file's mtime for nothing.
        const justOpened = lastSeenId !== round.id;
        lastSeenId = round.id;
        lastSeenUpdatedAt = round.updatedAt;
        if (justOpened) return;

        pending = { path: docPath, round };
        clearTimeout(timer);
        timer = setTimeout(flush, DEBOUNCE_MS);
    });

    return () => {
        unsubscribe();
        flush();
    };
}
