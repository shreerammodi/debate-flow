import { afterEach } from "vitest";

import { setFlowFs, type FlowFs, type FlowLocations } from "@/lib/persistence/flowFs";
import { basename, dedupeFilename, joinPath } from "@/lib/persistence/flowPaths";

export const HOME = "/home/test";
export const FLOWS_DIR = "/home/test/Documents/ebb";

export interface FakeFlowFs extends FlowFs {
    /** The virtual disk, keyed by absolute path. */
    files: Map<string, string>;
    /** What the next open picker returns; null cancels. */
    nextOpen: string | null;
    /** What the next save picker returns; null cancels. */
    nextSave: string | null;
    /** Paths passed to reveal, in order. */
    revealed: string[];
    /** Writes performed, in order, so tests can assert on autosave behavior. */
    writes: string[];
    /** Make the next write of any kind fail with this message. */
    failWrites: string | null;
}

/**
 * A filesystem the tests own outright.
 *
 * Every layer above the FlowFs port - the session, recents, the migration, the
 * start screen - runs against this unchanged, which is the reason the port
 * exists: none of that behavior needs Tauri IPC mocked to be tested.
 */
export function installFakeFlowFs(): FakeFlowFs {
    const fs: FakeFlowFs = {
        files: new Map(),
        nextOpen: null,
        nextSave: null,
        revealed: [],
        writes: [],
        failWrites: null,

        locations: () => Promise.resolve<FlowLocations>({ flowsDir: FLOWS_DIR, home: HOME }),

        pickOpenPath: () => Promise.resolve(fs.nextOpen),
        pickSavePath: () => Promise.resolve(fs.nextSave),

        createFlow: (dir, name, text) => {
            if (fs.failWrites) return Promise.reject(new Error(fs.failWrites));
            const taken = new Set(
                [...fs.files.keys()].filter((p) => p.startsWith(dir + "/")).map(basename),
            );
            const path = joinPath(dir, dedupeFilename(name, taken));
            fs.files.set(path, text);
            fs.writes.push(path);
            return Promise.resolve(path);
        },

        readFlow: (path) => Promise.resolve(fs.files.get(path) ?? null),

        writeFlow: (path, text) => {
            if (fs.failWrites) return Promise.reject(new Error(fs.failWrites));
            fs.files.set(path, text);
            fs.writes.push(path);
            return Promise.resolve();
        },

        readRecents: () => Promise.resolve(fs.files.get("/config/recents.json") ?? null),

        writeRecents: (text) => {
            fs.files.set("/config/recents.json", text);
            return Promise.resolve();
        },

        reveal: (path) => {
            fs.revealed.push(path);
            return Promise.resolve();
        },
    };

    setFlowFs(fs);
    afterEach(() => setFlowFs(null));
    return fs;
}
