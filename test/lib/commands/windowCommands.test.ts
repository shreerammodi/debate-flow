import { beforeEach, describe, expect, it, vi } from "vitest";

const invoke = vi.fn();
vi.mock("@tauri-apps/api/core", () => ({
    invoke: (cmd: string, args?: unknown) => invoke(cmd, args),
}));

let desktop = true;
vi.mock("@/lib/update/adapter", () => ({
    isDesktop: () => desktop,
}));

import { closeCurrentWindow, openNewWindow } from "@/lib/commands/windowCommands";

beforeEach(() => {
    invoke.mockReset();
    desktop = true;
});

describe("closeCurrentWindow", () => {
    it("closes through Rust, so the flush handshake runs first", async () => {
        await closeCurrentWindow();

        expect(invoke).toHaveBeenCalledWith("close_window", undefined);
    });

    it("asks no window manager on the web build", async () => {
        desktop = false;
        await closeCurrentWindow();

        expect(invoke).not.toHaveBeenCalled();
    });

    it("leaves the window open rather than throwing when the close fails", async () => {
        invoke.mockRejectedValueOnce(new Error("no such window"));

        await expect(closeCurrentWindow()).resolves.toBeUndefined();
    });
});

describe("openNewWindow", () => {
    it("opens through Rust, which owns window creation", async () => {
        await openNewWindow();

        expect(invoke).toHaveBeenCalledWith("new_window", undefined);
    });

    it("asks no window manager on the web build", async () => {
        desktop = false;
        await openNewWindow();

        expect(invoke).not.toHaveBeenCalled();
    });
});
