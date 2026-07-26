import { beforeEach, describe, expect, it } from "vitest";

import { resolveFlowsDir } from "@/lib/persistence/flowsDir";
import { useFlowStore } from "@/lib/store/useFlowStore";

import { FLOWS_DIR, installFakeFlowFs, type FakeFlowFs } from "../../support/fakeFlowFs";

let fs: FakeFlowFs;

beforeEach(() => {
    fs = installFakeFlowFs();
    useFlowStore.setState({ flowsDir: null });
});

describe("resolveFlowsDir", () => {
    it("falls back to the platform default when unset", async () => {
        expect(await resolveFlowsDir(fs)).toBe(FLOWS_DIR);
    });

    it("prefers the configured folder", async () => {
        useFlowStore.setState({ flowsDir: "/Volumes/usb/rounds" });
        expect(await resolveFlowsDir(fs)).toBe("/Volumes/usb/rounds");
    });

    it("treats a blank setting as unset rather than writing to nowhere", async () => {
        useFlowStore.setState({ flowsDir: "   " });
        expect(await resolveFlowsDir(fs)).toBe(FLOWS_DIR);
    });
});
