/**
 * In-memory adapter for `npm run dev` in a browser and for the test suite.
 *
 * Paths are virtual but behave like real ones, so everything above the port -
 * the session, recents, the migration - runs unchanged without Tauri. It is
 * backed by localStorage purely so a dev reload does not wipe the flow you were
 * looking at; this is a development surface, not a product one. The static
 * export exists only as Tauri's frontend and is not deployed anywhere.
 */

import type { FlowFs } from "./flowFs";
import { basename, dedupeFilename, joinPath } from "./flowPaths";

const STORE_KEY = "ebb-dev-flow-files";
const RECENTS_KEY = "ebb-dev-recents";
const HOME = "/home/dev";
const FLOWS_DIR = "/home/dev/Documents/ebb";

type Files = Record<string, string>;

function load(): Files {
    if (typeof localStorage === "undefined") return {};
    try {
        const raw = localStorage.getItem(STORE_KEY);
        return raw ? (JSON.parse(raw) as Files) : {};
    } catch {
        return {};
    }
}

function persist(files: Files): void {
    if (typeof localStorage === "undefined") return;
    try {
        localStorage.setItem(STORE_KEY, JSON.stringify(files));
    } catch {
        // A dev-only convenience; a full quota is not worth an error path.
    }
}

export function createFlowFs(): FlowFs {
    let files = load();

    /**
     * There is no picker without a native shell, so open and save fall back to
     * a prompt. Enough to exercise the real code paths in a browser.
     */
    const ask = (message: string, seed: string): string | null => {
        if (typeof window === "undefined") return null;
        const answer = window.prompt(message, seed);
        return answer?.trim() ? answer.trim() : null;
    };

    return {
        locations: () => Promise.resolve({ flowsDir: FLOWS_DIR, home: HOME }),

        pickOpenPath: () =>
            Promise.resolve(ask("Open which flow?", Object.keys(files)[0] ?? FLOWS_DIR)),

        pickSavePath: (suggested) =>
            Promise.resolve(ask("Save the flow as?", joinPath(FLOWS_DIR, basename(suggested)))),

        createFlow: (dir, name, text) => {
            const taken = new Set(
                Object.keys(files)
                    .filter((p) => p.startsWith(dir + "/"))
                    .map(basename),
            );
            const path = joinPath(dir, dedupeFilename(name, taken));
            files = { ...files, [path]: text };
            persist(files);
            return Promise.resolve(path);
        },

        readFlow: (path) => Promise.resolve(files[path] ?? null),

        writeFlow: (path, text) => {
            files = { ...files, [path]: text };
            persist(files);
            return Promise.resolve();
        },

        readRecents: () =>
            Promise.resolve(
                typeof localStorage === "undefined" ? null : localStorage.getItem(RECENTS_KEY),
            ),

        writeRecents: (text) => {
            if (typeof localStorage !== "undefined") localStorage.setItem(RECENTS_KEY, text);
            return Promise.resolve();
        },

        reveal: () => Promise.resolve(),
    };
}
