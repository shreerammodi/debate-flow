import { beforeEach, describe, expect, it, vi } from "vitest";

const listen = vi.fn();

vi.mock("@tauri-apps/api/event", () => ({
    listen: (...args: unknown[]) => listen(...args),
}));

vi.mock("@tauri-apps/api/webviewWindow", () => ({
    getCurrentWebviewWindow: () => ({ label: "win-3" }),
}));

import { listenHere } from "@/lib/windowEvents";

beforeEach(() => {
    listen.mockReset();
});

describe("listenHere", () => {
    // Tauri's targeting only holds if both halves name the label: `Emitter::emit`
    // reaches every webview whatever handle the shell called it on, and a
    // listener registered for the default `Any` target matches a narrowed emit
    // regardless. Dropping the target here silently un-narrows every
    // single-recipient event, and nothing else would notice.
    it("registers for this window rather than for any", async () => {
        listen.mockResolvedValue(() => {});
        await listenHere("collab:message", () => {});

        expect(listen).toHaveBeenCalledTimes(1);
        const [event, , options] = listen.mock.calls[0];
        expect(event).toBe("collab:message");
        expect(options).toEqual({ target: "win-3" });
    });

    it("hands the payload over without the event wrapper around it", async () => {
        listen.mockResolvedValue(() => {});
        const seen: unknown[] = [];
        await listenHere<{ connId: string }>("collab:peer", (p) => seen.push(p));

        const handler = listen.mock.calls[0][1] as (e: { payload: unknown }) => void;
        handler({ payload: { connId: "c1" } });

        expect(seen).toEqual([{ connId: "c1" }]);
    });

    it("stops listening through the closure it returns", async () => {
        const un = vi.fn();
        listen.mockResolvedValue(un);

        const stop = await listenHere("app:flush", () => {});
        expect(un).not.toHaveBeenCalled();
        stop();
        expect(un).toHaveBeenCalledTimes(1);
    });
});
