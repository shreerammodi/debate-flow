import { beforeEach, describe, expect, it } from "vitest";

import { askToShare, useCollabConsent } from "@/lib/store/useCollabConsent";
import { useFlowStore } from "@/lib/store/useFlowStore";

beforeEach(() => {
    (window as unknown as Record<string, unknown>).__TAURI_INTERNALS__ = {};
    useFlowStore.setState({ collabEnabled: false });
    useCollabConsent.getState().close();
});

describe("askToShare", () => {
    it("asks nothing when sharing is already on", async () => {
        useFlowStore.setState({ collabEnabled: true });
        await expect(askToShare()).resolves.toBe(true);
        expect(useCollabConsent.getState().open).toBe(false);
    });

    it("turns sharing on when the debater says so", async () => {
        const answer = askToShare();
        expect(useCollabConsent.getState().open).toBe(true);
        useCollabConsent.getState().answer(true);
        await expect(answer).resolves.toBe(true);
        expect(useFlowStore.getState().collabEnabled).toBe(true);
    });

    it("leaves the switch alone when they decline", async () => {
        const answer = askToShare();
        useCollabConsent.getState().answer(false);
        await expect(answer).resolves.toBe(false);
        expect(useFlowStore.getState().collabEnabled).toBe(false);
    });

    it("treats a dismissed dialog as no", async () => {
        const answer = askToShare();
        useCollabConsent.getState().close();
        await expect(answer).resolves.toBe(false);
        expect(useFlowStore.getState().collabEnabled).toBe(false);
    });

    it("settles a question still open when a second one is asked", async () => {
        const first = askToShare();
        const second = askToShare();
        useCollabConsent.getState().answer(true);
        await expect(first).resolves.toBe(false);
        await expect(second).resolves.toBe(true);
    });

    // A browser cannot bind an endpoint, so the question would offer something
    // that does not exist here.
    it("asks nothing off the desktop, and answers no", async () => {
        delete (window as unknown as Record<string, unknown>).__TAURI_INTERNALS__;
        await expect(askToShare()).resolves.toBe(false);
        expect(useCollabConsent.getState().open).toBe(false);
        expect(useFlowStore.getState().collabEnabled).toBe(false);
    });
});
