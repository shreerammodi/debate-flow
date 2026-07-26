import { beforeEach, describe, expect, it, vi } from "vitest";
import { createStore } from "zustand";

import { makeFlowRound, type FlowRound } from "@/lib/model/flow";
import { FLOW_FILE_VERSION, parseFlowFile, serializeFlow } from "@/lib/persistence/flowFile";
import {
    attachFlowAutosave,
    createFlowFile,
    pickFlowToOpen,
    readFlowAt,
    saveFlowAs,
    saveFlowNow,
} from "@/lib/persistence/flowSession";
import { loadRecents } from "@/lib/persistence/recents";

import { FLOWS_DIR, installFakeFlowFs, type FakeFlowFs } from "../../support/fakeFlowFs";

let fs: FakeFlowFs;

beforeEach(() => {
    fs = installFakeFlowFs();
});

describe("createFlowFile", () => {
    it("writes the flow into the flows folder and remembers it", async () => {
        const round = makeFlowRound({ event: "policy" });
        const path = await createFlowFile(round, fs);

        expect(path.startsWith(FLOWS_DIR + "/")).toBe(true);
        expect(parseFlowFile(fs.files.get(path)!).id).toBe(round.id);
        expect((await loadRecents(fs)).map((r) => r.path)).toEqual([path]);
    });

    it("never overwrites an existing flow", async () => {
        const createdAt = new Date(2026, 6, 25).getTime();
        const first = await createFlowFile({ ...makeFlowRound({}), createdAt }, fs);
        const second = await createFlowFile({ ...makeFlowRound({}), createdAt }, fs);

        expect(second).not.toBe(first);
        expect(fs.files.size).toBeGreaterThanOrEqual(2);
    });
});

describe("readFlowAt", () => {
    it("returns null for a file that is gone, which is an ordinary outcome", async () => {
        expect(await readFlowAt("/nowhere.ebb", fs)).toBeNull();
    });

    it("throws for a file that exists but will not parse", async () => {
        fs.files.set("/a.ebb", "{ truncated");
        await expect(readFlowAt("/a.ebb", fs)).rejects.toThrow(/not valid JSON/);
    });
});

describe("pickFlowToOpen", () => {
    it("returns null when the picker is cancelled", async () => {
        fs.nextOpen = null;
        expect(await pickFlowToOpen(fs)).toBeNull();
    });

    it("opens a .ebb in place and records it", async () => {
        const round = makeFlowRound({});
        fs.files.set("/elsewhere/r.ebb", serializeFlow(round));
        fs.nextOpen = "/elsewhere/r.ebb";

        expect(await pickFlowToOpen(fs)).toBe("/elsewhere/r.ebb");
        expect((await loadRecents(fs)).map((r) => r.path)).toEqual(["/elsewhere/r.ebb"]);
        // Opening must not copy the file into the flows folder.
        expect([...fs.files.keys()]).not.toContainEqual(expect.stringContaining(FLOWS_DIR));
    });

    it("materializes a legacy backup into one file per round", async () => {
        const rounds = [makeFlowRound({ event: "policy" }), makeFlowRound({ event: "ld" })];
        fs.files.set(
            "/downloads/backup.json",
            JSON.stringify({ version: FLOW_FILE_VERSION, kind: "backup", rounds }),
        );
        fs.nextOpen = "/downloads/backup.json";

        const opened = await pickFlowToOpen(fs);
        const written = [...fs.files.keys()].filter((p) => p.startsWith(FLOWS_DIR));

        expect(written).toHaveLength(2);
        expect(opened).toBe(written[0]);
        expect(parseFlowFile(fs.files.get(opened!)!).event).toBe("policy");
    });
});

describe("saveFlowAs", () => {
    it("writes to the picked path and remembers it", async () => {
        const round = makeFlowRound({});
        fs.nextSave = "/elsewhere/copy.ebb";

        expect(await saveFlowAs(round, fs)).toBe("/elsewhere/copy.ebb");
        expect(parseFlowFile(fs.files.get("/elsewhere/copy.ebb")!).id).toBe(round.id);
        expect((await loadRecents(fs)).map((r) => r.path)).toEqual(["/elsewhere/copy.ebb"]);
    });

    it("writes nothing when cancelled", async () => {
        fs.nextSave = null;
        expect(await saveFlowAs(makeFlowRound({}), fs)).toBeNull();
        expect(fs.writes).toEqual([]);
    });
});

describe("saveFlowNow", () => {
    it("reports saving then saved", async () => {
        const status = vi.fn();
        await saveFlowNow("/a.ebb", makeFlowRound({}), status);
        expect(status.mock.calls.flat()).toEqual(["saving", "saved"]);
    });

    it("reports an error instead of throwing at the caller", async () => {
        const status = vi.fn();
        fs.failWrites = "disk full";
        await saveFlowNow("/a.ebb", makeFlowRound({}), status);
        expect(status.mock.calls.flat()).toEqual(["saving", "error"]);
    });
});

describe("attachFlowAutosave", () => {
    interface DocState {
        round: FlowRound | null;
        docPath: string | null;
    }

    /**
     * The store is empty until a flow is loaded, exactly as in the app: zustand
     * only notifies on setState, so the notification that carries the round in
     * is the "just opened" one autosave must not treat as an edit.
     */
    function openIn(round: FlowRound, docPath = "/a.ebb") {
        const s = createStore<DocState>()(() => ({ round: null, docPath: null }));
        const stop = attachFlowAutosave(s);
        s.setState({ round, docPath });
        return { store: s, stop };
    }

    beforeEach(() => vi.useFakeTimers());

    it("does not rewrite a flow just because it was opened", async () => {
        const { stop } = openIn(makeFlowRound({}));
        await vi.advanceTimersByTimeAsync(600);

        expect(fs.writes).toEqual([]);
        stop();
    });

    it("debounces a burst of edits into one write", async () => {
        const round = makeFlowRound({});
        const { store: s, stop } = openIn(round);

        for (let i = 1; i <= 5; i++) {
            s.setState({ round: { ...round, updatedAt: round.updatedAt + i } });
        }
        expect(fs.writes).toEqual([]);

        await vi.advanceTimersByTimeAsync(600);
        expect(fs.writes).toEqual(["/a.ebb"]);
        expect(parseFlowFile(fs.files.get("/a.ebb")!).updatedAt).toBe(round.updatedAt + 5);
        stop();
    });

    it("flushes on teardown, so navigating away never drops the last edit", async () => {
        const round = makeFlowRound({});
        const { store: s, stop } = openIn(round);

        s.setState({ round: { ...round, updatedAt: round.updatedAt + 1 } });
        stop();
        await vi.advanceTimersByTimeAsync(0);

        expect(fs.writes).toEqual(["/a.ebb"]);
    });

    it("ignores a store change that did not touch the round", async () => {
        const round = makeFlowRound({});
        const { store: s, stop } = openIn(round);

        s.setState({ round: { ...round } });
        await vi.advanceTimersByTimeAsync(600);

        expect(fs.writes).toEqual([]);
        stop();
    });

    it("writes nothing while no file is open", async () => {
        const round = makeFlowRound({});
        const s = createStore<DocState>()(() => ({ round: null, docPath: null }));
        const stop = attachFlowAutosave(s);

        s.setState({ round, docPath: null });
        s.setState({ round: { ...round, updatedAt: round.updatedAt + 1 } });
        await vi.advanceTimersByTimeAsync(600);

        expect(fs.writes).toEqual([]);
        stop();
    });

    it("follows the round to a new path after Save As", async () => {
        const round = makeFlowRound({});
        const { store: s, stop } = openIn(round);

        s.setState({ round: { ...round, updatedAt: round.updatedAt + 1 }, docPath: "/b.ebb" });
        await vi.advanceTimersByTimeAsync(600);

        expect(fs.writes).toEqual(["/b.ebb"]);
        stop();
    });
});
