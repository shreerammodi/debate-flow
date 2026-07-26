import { render, waitFor } from "@testing-library/react";
import { beforeEach, describe, expect, it, vi } from "vitest";

const invoke = vi.fn();
let flushHandler: (() => void) | null = null;
const unlisten = vi.fn();

vi.mock("@tauri-apps/api/core", () => ({ invoke: (...a: unknown[]) => invoke(...a) }));
vi.mock("@tauri-apps/api/event", () => ({
    listen: (name: string, cb: () => void) => {
        if (name === "app:flush") flushHandler = cb;
        return Promise.resolve(unlisten);
    },
}));

const saveOpenFlow = vi.fn();
vi.mock("@/lib/commands/fileCommands", () => ({ saveOpenFlow: () => saveOpenFlow() }));

const toastError = vi.fn();
vi.mock("sonner", () => ({ toast: { error: (...a: unknown[]) => toastError(...a) } }));

vi.mock("@/lib/update/adapter", () => ({ isDesktop: () => true }));

import QuitGuard from "@/components/QuitGuard";

beforeEach(() => {
    invoke.mockReset();
    saveOpenFlow.mockReset();
    toastError.mockReset();
    flushHandler = null;
});

/** Mount, wait for the listener, then act as the shell asking for a flush. */
async function requestQuit() {
    render(<QuitGuard />);
    await waitFor(() => expect(flushHandler).not.toBeNull());
    flushHandler?.();
}

describe("QuitGuard", () => {
    it("writes the open flow before letting the app exit", async () => {
        saveOpenFlow.mockResolvedValue(true);

        await requestQuit();

        await waitFor(() => expect(invoke).toHaveBeenCalledWith("finish_quit", { saved: true }));
        expect(saveOpenFlow).toHaveBeenCalled();
    });

    it("cancels the exit when the flow could not be written", async () => {
        // Quitting with a full disk must keep the window, not take the round
        // down with the process.
        saveOpenFlow.mockResolvedValue(false);

        await requestQuit();

        await waitFor(() => expect(invoke).toHaveBeenCalledWith("finish_quit", { saved: false }));
        expect(toastError).toHaveBeenCalled();
    });

    it("treats a thrown save as a failure rather than exiting anyway", async () => {
        saveOpenFlow.mockRejectedValue(new Error("boom"));

        await requestQuit();

        await waitFor(() => expect(invoke).toHaveBeenCalledWith("finish_quit", { saved: false }));
    });
});
