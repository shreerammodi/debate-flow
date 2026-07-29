/**
 * The id Settings shows, and what asking for it costs.
 *
 * It used to arrive from the idle listener, so seeing your own id meant an
 * endpoint was bound. It comes off the identity file now, which is the whole
 * point: a debater can hand a partner their id with nothing on the network.
 */

import { beforeEach, describe, expect, it, vi } from "vitest";

import { clearMyEndpointId, myEndpointId } from "@/lib/collab/identity";

const invoke = vi.hoisted(() => vi.fn());

vi.mock("@tauri-apps/api/core", () => ({ invoke }));

/** isDesktop() reads this, the way the rest of the suite fakes a shell. */
function onDesktop(yes: boolean): void {
    if (yes) (window as unknown as Record<string, unknown>).__TAURI_INTERNALS__ = {};
    else delete (window as unknown as Record<string, unknown>).__TAURI_INTERNALS__;
}

const ID = "e".repeat(64);

beforeEach(() => {
    invoke.mockReset();
    clearMyEndpointId();
    onDesktop(true);
});

describe("myEndpointId", () => {
    it("reads the identity off the shell, binding nothing", async () => {
        invoke.mockResolvedValue(ID);
        expect(await myEndpointId()).toBe(ID);
        expect(invoke).toHaveBeenCalledWith("collab_endpoint_id");
        expect(invoke).toHaveBeenCalledTimes(1);
    });

    it("asks once, because the identity cannot change while the app is open", async () => {
        invoke.mockResolvedValue(ID);
        await myEndpointId();
        await myEndpointId();
        expect(invoke).toHaveBeenCalledTimes(1);
    });

    it("reports no id rather than failing when the shell cannot say", async () => {
        invoke.mockRejectedValue(new Error("no identity file"));
        expect(await myEndpointId()).toBe("");
    });

    it("asks nothing on web, where there is no shell to ask", async () => {
        onDesktop(false);
        expect(await myEndpointId()).toBe("");
        expect(invoke).not.toHaveBeenCalled();
    });
});
