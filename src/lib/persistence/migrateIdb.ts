/**
 * One-time sweep of the flows that used to live in IndexedDB.
 *
 * Ebb kept every round in a single-table Dexie database before flows became
 * files. That data is a user's tournament history, so the sweep runs itself: it
 * writes each round into the flows folder, reads every file back and parses it,
 * and only then deletes the database. Any failure leaves the source untouched,
 * because a half-finished migration that has already dropped its input is the
 * one outcome worse than not migrating at all.
 *
 * The raw IndexedDB API is used rather than Dexie so the dependency can go.
 * Soft-deleted rounds are written to a trash subfolder instead of discarded:
 * the trash concept is gone, but the rounds that were in it are still the
 * user's.
 */

import { normalizeFlow, type FlowRound } from "@/lib/model/flow";

import { parseFlowFile, serializeFlow } from "./flowFile";
import { getFlowFs, type FlowFs } from "./flowFs";
import { joinPath, suggestFilename } from "./flowPaths";
import { loadRecents, promoteRecent, saveRecents } from "./recents";

const DB_NAME = "ebbflow";
const STORE = "flows";
const DONE_KEY = "ebb-idb-migrated";

export interface MigrationReport {
    /** Live rounds written into the flows folder. */
    moved: number;
    /** Soft-deleted rounds written into its trash subfolder. */
    trashed: number;
    flowsDir: string;
}

function openDb(): Promise<IDBDatabase | null> {
    const { promise, resolve } = Promise.withResolvers<IDBDatabase | null>();
    // Opening creates an empty database when none exists, which reads as "no
    // flows to migrate" and is deleted along with a real one.
    const req = indexedDB.open(DB_NAME);
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => resolve(null);
    req.onblocked = () => resolve(null);
    return promise;
}

function readAll(db: IDBDatabase): Promise<unknown[]> {
    if (!db.objectStoreNames.contains(STORE)) return Promise.resolve([]);
    const { promise, resolve, reject } = Promise.withResolvers<unknown[]>();
    const req = db.transaction(STORE, "readonly").objectStore(STORE).getAll();
    req.onsuccess = () => resolve(req.result as unknown[]);
    req.onerror = () => reject(req.error ?? new Error("Could not read the old flows"));
    return promise;
}

function deleteDb(): Promise<void> {
    const { promise, resolve } = Promise.withResolvers<void>();
    const req = indexedDB.deleteDatabase(DB_NAME);
    // A blocked or failed deletion is not worth surfacing: the files are
    // already written and verified, and the marker stops a second sweep.
    req.onsuccess = () => resolve();
    req.onerror = () => resolve();
    req.onblocked = () => resolve();
    return promise;
}

/**
 * Move every stored round into files. Resolves to null when there is nothing to
 * migrate, which is the case for every launch after the first.
 */
export async function migrateFromIndexedDb(fs?: FlowFs): Promise<MigrationReport | null> {
    if (typeof indexedDB === "undefined") return null;
    if (typeof localStorage !== "undefined" && localStorage.getItem(DONE_KEY)) return null;

    const db = await openDb();
    if (!db) return null;

    let records: unknown[];
    try {
        records = await readAll(db);
    } finally {
        db.close();
    }

    if (!records.length) {
        await deleteDb();
        localStorage?.setItem(DONE_KEY, "1");
        return null;
    }

    const io = fs ?? (await getFlowFs());
    const { flowsDir } = await io.locations();
    const trashDir = joinPath(flowsDir, "trash");

    const written: { path: string; live: boolean }[] = [];
    for (const record of records) {
        const trashed =
            typeof record === "object" &&
            record !== null &&
            "deletedAt" in record &&
            record.deletedAt != null;

        // normalizeFlow drops deletedAt; the subfolder carries that fact now.
        const round = normalizeFlow(record as FlowRound);
        const path = await io.createFlow(
            trashed ? trashDir : flowsDir,
            suggestFilename(round),
            serializeFlow(round),
        );
        written.push({ path, live: !trashed });
    }

    // Read every file back before dropping the source. A write that reported
    // success but produced something unparseable is exactly the failure this
    // guards, and it is only detectable from the far side.
    for (const { path } of written) {
        const text = await io.readFlow(path);
        if (text === null) throw new Error(`Migration wrote ${path} but could not read it back`);
        parseFlowFile(text);
    }

    await deleteDb();
    localStorage?.setItem(DONE_KEY, "1");

    // Seed the start screen oldest-first, so the most recently updated round
    // ends up at the top of the list.
    const live = written.filter((w) => w.live);
    let recents = await loadRecents(io);
    for (const { path } of live) recents = promoteRecent(recents, path, Date.now());
    await saveRecents(io, recents);

    return {
        moved: live.length,
        trashed: written.length - live.length,
        flowsDir,
    };
}
