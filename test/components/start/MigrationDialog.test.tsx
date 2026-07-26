/**
 * IMPORTANT: fake-indexeddb/auto must be imported first so it polyfills the
 * global indexedDB before the legacy check reaches for it.
 */
import "fake-indexeddb/auto";
import { render, screen, waitFor } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { beforeEach, describe, expect, it, vi } from "vitest";

import MigrationDialog from "@/components/start/MigrationDialog";
import { makeFlowRound } from "@/lib/model/flow";
import { useFlowStore } from "@/lib/store/useFlowStore";

import { FLOWS_DIR, installFakeFlowFs, type FakeFlowFs } from "../../support/fakeFlowFs";

const DB_NAME = "ebbflow";
const DONE_KEY = "ebb-idb-migrated";

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

function dropDb(): Promise<void> {
    const { promise, resolve } = Promise.withResolvers<void>();
    const req = indexedDB.deleteDatabase(DB_NAME);
    req.onsuccess = () => resolve();
    req.onerror = () => resolve();
    req.onblocked = () => resolve();
    return promise;
}

let fs: FakeFlowFs;

beforeEach(async () => {
    fs = installFakeFlowFs();
    localStorage.removeItem(DONE_KEY);
    useFlowStore.setState({ flowsDir: null });
    await dropDb();
});

describe("MigrationDialog", () => {
    it("stays out of the way when there is nothing to move", async () => {
        render(<MigrationDialog onMigrated={vi.fn()} />);
        await waitFor(() => expect(localStorage.getItem(DONE_KEY)).toBe("1"));
        expect(screen.queryByTestId("migration-dialog")).not.toBeInTheDocument();
    });

    it("asks before moving anything, and moves nothing until told to", async () => {
        await seedDb([makeFlowRound({}), makeFlowRound({})]);

        render(<MigrationDialog onMigrated={vi.fn()} />);

        expect(await screen.findByTestId("migration-dialog")).toHaveTextContent(
            "Move 2 flows into files?",
        );
        // The whole point: consent first. Nothing on disk yet.
        expect(fs.writes).toEqual([]);
    });

    it("shows where the flows will land", async () => {
        await seedDb([makeFlowRound({})]);
        render(<MigrationDialog onMigrated={vi.fn()} />);
        expect(await screen.findByTestId("migration-target")).toHaveTextContent(FLOWS_DIR);
    });

    it("moves the flows and reports back when confirmed", async () => {
        await seedDb([makeFlowRound({}), makeFlowRound({})]);
        const onMigrated = vi.fn();
        render(<MigrationDialog onMigrated={onMigrated} />);
        await screen.findByTestId("migration-dialog");

        await userEvent.click(screen.getByTestId("migration-move"));

        await waitFor(() => expect(onMigrated).toHaveBeenCalled());
        expect([...fs.files.keys()].filter((p) => p.endsWith(".ebb"))).toHaveLength(2);
        expect(screen.queryByTestId("migration-dialog")).not.toBeInTheDocument();
    });

    it("leaves the old flows alone on Not now, and asks again next launch", async () => {
        await seedDb([makeFlowRound({})]);
        const { unmount } = render(<MigrationDialog onMigrated={vi.fn()} />);
        await screen.findByTestId("migration-dialog");

        await userEvent.click(screen.getByTestId("migration-later"));
        expect(screen.queryByTestId("migration-dialog")).not.toBeInTheDocument();
        expect(fs.writes).toEqual([]);
        // Declining must not settle the marker, or the rounds would be stranded.
        expect(localStorage.getItem(DONE_KEY)).toBeNull();

        unmount();
        render(<MigrationDialog onMigrated={vi.fn()} />);
        expect(await screen.findByTestId("migration-dialog")).toBeInTheDocument();
    });

    it("migrates into a folder chosen from the prompt", async () => {
        await seedDb([makeFlowRound({})]);
        fs.nextDirectory = "/Volumes/usb/rounds";
        render(<MigrationDialog onMigrated={vi.fn()} />);
        await screen.findByTestId("migration-dialog");

        await userEvent.click(screen.getByTestId("migration-choose-folder"));
        await waitFor(() =>
            expect(screen.getByTestId("migration-target")).toHaveTextContent("/Volumes/usb/rounds"),
        );
        await userEvent.click(screen.getByTestId("migration-move"));

        await waitFor(() =>
            expect([...fs.files.keys()].some((p) => p.startsWith("/Volumes/usb/rounds/"))).toBe(
                true,
            ),
        );
        // The choice sticks as the setting, so new flows go there too.
        expect(useFlowStore.getState().flowsDir).toBe("/Volumes/usb/rounds");
    });
});
