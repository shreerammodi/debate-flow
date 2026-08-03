import { beforeEach, describe, expect, it, vi } from "vitest";

import {
    closeShareSheet,
    openShareSheet,
    showShareCode,
    showShareFailure,
    showShareGuest,
    useShareSheet,
} from "@/lib/store/useShareSheet";

beforeEach(() => {
    closeShareSheet();
});

describe("the share sheet", () => {
    it("opens on Getting ready, with no code on it", () => {
        openShareSheet("editor", "");
        const s = useShareSheet.getState();
        expect(s.open).toBe(true);
        expect(s.screen).toBe("ready");
        expect(s.code).toBe("");
    });

    it("carries the warning a debater is owed before a code exists", () => {
        openShareSheet("editor", "Relaying is off, so this code only works on the same wifi.");
        expect(useShareSheet.getState().warning).toBe(
            "Relaying is off, so this code only works on the same wifi.",
        );
    });

    it("shows the code once there is one", () => {
        openShareSheet("editor", "");
        showShareCode("K7QM3XPV", async () => {});
        const s = useShareSheet.getState();
        expect(s.screen).toBe("code");
        expect(s.code).toBe("K7QM3XPV");
    });

    it("names the guest who arrived", () => {
        openShareSheet("editor", "");
        showShareCode("K7QM3XPV", async () => {});
        showShareGuest("Sam");
        const s = useShareSheet.getState();
        expect(s.screen).toBe("joined");
        expect(s.guest).toBe("Sam");
    });

    it("shows the reason instead of a code when there is none", () => {
        openShareSheet("editor", "");
        showShareFailure("Could not reach the relay for that code");
        const s = useShareSheet.getState();
        expect(s.screen).toBe("failed");
        expect(s.message).toBe("Could not reach the relay for that code");
        expect(s.code).toBe("");
    });

    it("kills the code when the sheet closes", () => {
        const stop = vi.fn(async () => {});
        openShareSheet("editor", "");
        showShareCode("K7QM3XPV", stop);
        closeShareSheet();
        expect(stop).toHaveBeenCalledOnce();
        expect(useShareSheet.getState().open).toBe(false);
    });

    it("kills the old code when a second share opens over it", () => {
        const stop = vi.fn(async () => {});
        openShareSheet("editor", "");
        showShareCode("K7QM3XPV", stop);
        openShareSheet("viewer", "");
        expect(stop).toHaveBeenCalledOnce();
        expect(useShareSheet.getState().code).toBe("");
    });

    it("kills the code when the sheet gives up on making another", () => {
        const stop = vi.fn(async () => {});
        openShareSheet("editor", "");
        showShareCode("K7QM3XPV", stop);
        showShareFailure("Could not reach the relay for that code");
        expect(stop).toHaveBeenCalledOnce();
    });

    it("kills one code once, however many times the sheet is closed", () => {
        const stop = vi.fn(async () => {});
        openShareSheet("editor", "");
        showShareCode("K7QM3XPV", stop);
        closeShareSheet();
        closeShareSheet();
        expect(stop).toHaveBeenCalledOnce();
    });

    it("swallows a teardown that failed, because nobody is waiting on it", () => {
        openShareSheet("editor", "");
        showShareCode("K7QM3XPV", async () => {
            throw new Error("the endpoint refused to stop");
        });
        expect(() => closeShareSheet()).not.toThrow();
    });
});
