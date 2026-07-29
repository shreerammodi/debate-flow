/**
 * IMPORTANT: fake-indexeddb/auto must be imported first so it polyfills the
 * global indexedDB before the migration reaches for it. This is the only suite
 * that still needs it - everything else runs on the fake filesystem.
 */
import "fake-indexeddb/auto";
import { afterEach, beforeEach, describe, expect, it } from "vitest";

import { makeFlowRound, type FlowRound } from "@/lib/model/flow";
import { parseFlowFile } from "@/lib/persistence/flowFile";
import { migrateFromIndexedDb } from "@/lib/persistence/migrateIdb";
import { loadRecents } from "@/lib/persistence/recents";

import { FLOWS_DIR, installFakeFlowFs, type FakeFlowFs } from "../../support/fakeFlowFs";

const DB_NAME = "ebbflow";
const DONE_KEY = "ebb-idb-migrated";

/** Build the single-table database the pre-file builds of ebb wrote. */
function seedDb(rounds: unknown[]): Promise<void> {
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

function dbExists(): Promise<boolean> {
    return indexedDB.databases().then((dbs) => dbs.some((d) => d.name === DB_NAME));
}

let fs: FakeFlowFs;

beforeEach(async () => {
    fs = installFakeFlowFs();
    localStorage.removeItem(DONE_KEY);
    const { promise, resolve } = Promise.withResolvers<void>();
    const req = indexedDB.deleteDatabase(DB_NAME);
    req.onsuccess = () => resolve();
    req.onerror = () => resolve();
    req.onblocked = () => resolve();
    await promise;
});

afterEach(() => localStorage.removeItem(DONE_KEY));

describe("migrateFromIndexedDb", () => {
    it("writes every stored round into the flows folder", async () => {
        const live = [makeFlowRound({ event: "policy" }), makeFlowRound({ event: "ld" })];
        await seedDb(live);

        const report = await migrateFromIndexedDb(FLOWS_DIR, fs);

        expect(report).toMatchObject({ moved: 2, trashed: 0, flowsDir: FLOWS_DIR });
        const written = [...fs.files.keys()].filter((p) => p.endsWith(".ebb"));
        expect(written).toHaveLength(2);
        expect(written.every((p) => p.startsWith(FLOWS_DIR + "/"))).toBe(true);
    });

    it("preserves each round's identity and content", async () => {
        const round = makeFlowRound({ event: "pf", firstSide: "neg" });
        round.sheets[1].data = [["framework"]];
        await seedDb([round]);

        await migrateFromIndexedDb(FLOWS_DIR, fs);

        const [path] = [...fs.files.keys()].filter((p) => p.endsWith(".ebb"));
        const migrated = parseFlowFile(fs.files.get(path)!);
        expect(migrated.id).toBe(round.id);
        expect(migrated.firstSide).toBe("neg");
        expect(migrated.sheets[1].data).toEqual([["framework"]]);
    });

    it("keeps trashed rounds in a trash subfolder rather than discarding them", async () => {
        const trashed: FlowRound & { deletedAt: number } = {
            ...makeFlowRound({}),
            deletedAt: Date.now(),
        };
        await seedDb([makeFlowRound({}), trashed]);

        const report = await migrateFromIndexedDb(FLOWS_DIR, fs);

        expect(report).toMatchObject({ moved: 1, trashed: 1 });
        const inTrash = [...fs.files.keys()].filter((p) => p.startsWith(`${FLOWS_DIR}/trash/`));
        expect(inTrash).toHaveLength(1);
        // The trash concept is gone, so the migrated file must not carry it.
        expect(parseFlowFile(fs.files.get(inTrash[0])!)).not.toHaveProperty("deletedAt");
    });

    it("seeds recents with the live flows only", async () => {
        await seedDb([makeFlowRound({}), { ...makeFlowRound({}), deletedAt: 1 }]);
        await migrateFromIndexedDb(FLOWS_DIR, fs);

        const recents = await loadRecents(fs);
        expect(recents).toHaveLength(1);
        expect(recents[0].path.startsWith(`${FLOWS_DIR}/trash/`)).toBe(false);
    });

    it("deletes the database once every file has been read back", async () => {
        await seedDb([makeFlowRound({})]);
        await migrateFromIndexedDb(FLOWS_DIR, fs);
        expect(await dbExists()).toBe(false);
    });

    it("keeps the database when a write fails, rather than losing the source", async () => {
        await seedDb([makeFlowRound({})]);
        fs.failWrites = "disk full";

        await expect(migrateFromIndexedDb(FLOWS_DIR, fs)).rejects.toThrow(/disk full/);
        expect(await dbExists()).toBe(true);
        expect(localStorage.getItem(DONE_KEY)).toBeNull();
    });

    it("keeps the database when a record is not a round this build can render", async () => {
        // The node model that predates version 3. normalizeFlow would fill it out
        // into a structurally valid, empty round, so the read-back guard cannot
        // catch it: the content is dropped before the write.
        await seedDb([
            { id: "legacy", createdAt: 1, updatedAt: 2, nodes: [{ text: "perm" }] },
            makeFlowRound({}),
        ]);

        await expect(migrateFromIndexedDb(FLOWS_DIR, fs)).rejects.toThrow(/record\.scouting/);
        expect(await dbExists()).toBe(true);
        expect([...fs.files.keys()].filter((p) => p.endsWith(".ebb"))).toEqual([]);
        expect(localStorage.getItem(DONE_KEY)).toBeNull();
    });

    it("runs once; a second launch finds nothing to do", async () => {
        await seedDb([makeFlowRound({})]);
        expect(await migrateFromIndexedDb(FLOWS_DIR, fs)).not.toBeNull();

        const before = fs.files.size;
        expect(await migrateFromIndexedDb(FLOWS_DIR, fs)).toBeNull();
        expect(fs.files.size).toBe(before);
    });

    it("reports nothing when there was never a database", async () => {
        expect(await migrateFromIndexedDb(FLOWS_DIR, fs)).toBeNull();
        expect([...fs.files.keys()].filter((p) => p.endsWith(".ebb"))).toEqual([]);
    });
});
