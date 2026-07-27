import { beforeEach, describe, expect, it } from "vitest";

import { getSidecarFs, setSidecarFs } from "@/lib/collab/sidecarFs";
import { createSidecarFs } from "@/lib/collab/sidecarFsMemory";

beforeEach(() => {
    setSidecarFs(null);
});

describe("sidecarFsMemory", () => {
    it("reads back what it wrote, per round", async () => {
        const fs = createSidecarFs();
        await fs.write("round_a", "one");
        await fs.write("round_b", "two");
        expect(await fs.read("round_a")).toBe("one");
        expect(await fs.read("round_b")).toBe("two");
    });

    it("reports a round it has never seen as absent, not as an error", async () => {
        expect(await createSidecarFs().read("round_missing")).toBeNull();
    });

    it("overwrites in place", async () => {
        const fs = createSidecarFs();
        await fs.write("round_a", "one");
        await fs.write("round_a", "two");
        expect(await fs.read("round_a")).toBe("two");
    });
});

describe("getSidecarFs", () => {
    it("resolves the memory adapter off the desktop", async () => {
        expect(await getSidecarFs()).toBeDefined();
    });

    it("hands back the fixture a test installs", async () => {
        const fixture = { read: async () => "fixed", write: async () => {} };
        setSidecarFs(fixture);
        expect(await getSidecarFs()).toBe(fixture);
    });
});
