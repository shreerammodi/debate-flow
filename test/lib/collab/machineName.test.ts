/**
 * The name a session broadcasts, and where it comes from.
 *
 * The hostname is never written into the config file, so the resolution order
 * is the whole contract: what the debater typed, else what the machine is
 * called, else nothing.
 */

import { beforeEach, describe, expect, it, vi } from "vitest";

import { broadcastName, clearMachineName, machineName } from "@/lib/collab/machineName";
import { useFlowStore } from "@/lib/store/useFlowStore";

const invoke = vi.hoisted(() => vi.fn());

vi.mock("@tauri-apps/api/core", () => ({ invoke }));

/** isDesktop() reads this, the way the rest of the suite fakes a shell. */
function onDesktop(yes: boolean): void {
    if (yes) (window as unknown as Record<string, unknown>).__TAURI_INTERNALS__ = {};
    else delete (window as unknown as Record<string, unknown>).__TAURI_INTERNALS__;
}

beforeEach(() => {
    invoke.mockReset();
    clearMachineName();
    onDesktop(true);
    useFlowStore.setState({ collabName: "" });
});

describe("machineName", () => {
    it("asks the shell for the hostname", async () => {
        invoke.mockResolvedValue("smodi-mbp");
        expect(await machineName()).toBe("smodi-mbp");
        expect(invoke).toHaveBeenCalledWith("machine_name");
    });

    it("asks once, because a hostname cannot change while the app is open", async () => {
        invoke.mockResolvedValue("smodi-mbp");
        await machineName();
        await machineName();
        expect(invoke).toHaveBeenCalledTimes(1);
    });

    it("reports no name rather than failing when the shell cannot say", async () => {
        invoke.mockRejectedValue(new Error("no hostname binary"));
        expect(await machineName()).toBe("");
    });

    it("asks nothing on web, where there is no shell to ask", async () => {
        onDesktop(false);
        expect(await machineName()).toBe("");
        expect(invoke).not.toHaveBeenCalled();
    });
});

describe("broadcastName", () => {
    it("falls back to the machine's own name", async () => {
        invoke.mockResolvedValue("smodi-mbp");
        expect(await broadcastName()).toBe("smodi-mbp");
    });

    it("prefers the name the debater typed", async () => {
        invoke.mockResolvedValue("smodi-mbp");
        useFlowStore.setState({ collabName: "Rin" });
        expect(await broadcastName()).toBe("Rin");
    });

    it("treats a whitespace name as none at all", async () => {
        invoke.mockResolvedValue("smodi-mbp");
        useFlowStore.setState({ collabName: "   " });
        expect(await broadcastName()).toBe("smodi-mbp");
    });

    it("greets with no name when there is neither", async () => {
        invoke.mockResolvedValue("");
        expect(await broadcastName()).toBe("");
    });
});
