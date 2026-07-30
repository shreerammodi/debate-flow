/**
 * The legacy IndexedDB store the pre-file builds of ebb wrote.
 *
 * Only the migration and the dialog that offers it still care about it, and
 * both need the same database seeded the same way. Callers must import
 * "fake-indexeddb/auto" before anything that reaches for `indexedDB`.
 */

export const DB_NAME = "ebbflow";
export const DONE_KEY = "ebb-idb-migrated";

/** Build the single-table database the pre-file builds of ebb wrote. */
export function seedDb(rounds: unknown[]): Promise<void> {
    const { promise, resolve, reject } = Promise.withResolvers<void>();
    const req = indexedDB.open(DB_NAME, 1);
    req.onupgradeneeded = () => req.result.createObjectStore("flows", { keyPath: "id" });
    req.onerror = () => reject(req.error);
    req.onsuccess = () => {
        const db = req.result;
        const tx = db.transaction("flows", "readwrite");
        for (const round of rounds) tx.objectStore("flows").put(round);
        tx.oncomplete = () => {
            db.close();
            resolve();
        };
        tx.onerror = () => reject(tx.error);
    };
    return promise;
}

/** Deletes the database, resolving whether or not it was there to delete. */
export function dropDb(): Promise<void> {
    const { promise, resolve } = Promise.withResolvers<void>();
    const req = indexedDB.deleteDatabase(DB_NAME);
    req.onsuccess = () => resolve();
    req.onerror = () => resolve();
    req.onblocked = () => resolve();
    return promise;
}

export function dbExists(): Promise<boolean> {
    return indexedDB.databases().then((dbs) => dbs.some((d) => d.name === DB_NAME));
}
