import { describe, expect, it } from "vitest";

import { LOCK_TTL_MS, type Lock } from "@/lib/collab/presence";
import { LOCK_CLASS, lockClassFor, lockLabel } from "@/lib/grid/lockDecor";

const held = (
    endpointId: string,
    col: number,
    row: number,
    heldAt = 1_000,
    sheetId = "s1",
): Lock => ({
    endpointId,
    sheetId,
    col,
    row,
    heldAt,
});

const NAMES: Record<string, string> = { sam: "Sam", kim: "Kim" };
const nameOf = (endpointId: string) => NAMES[endpointId] ?? endpointId;

describe("lockClassFor", () => {
    it("marks the cell a peer holds", () => {
        expect(lockClassFor([held("sam", 2, 4)], "s1", 2, 4, 1_000)).toBe(LOCK_CLASS);
    });

    it("leaves a neighbouring cell unmarked", () => {
        const locks = [held("sam", 2, 4)];
        expect(lockClassFor(locks, "s1", 2, 5, 1_000)).toBeNull();
        expect(lockClassFor(locks, "s1", 3, 4, 1_000)).toBeNull();
    });

    it("leaves the same coordinates on another sheet unmarked", () => {
        expect(lockClassFor([held("sam", 2, 4)], "s2", 2, 4, 1_000)).toBeNull();
    });

    it("marks nothing at all when no peer holds anything", () => {
        expect(lockClassFor([], "s1", 0, 0, 1_000)).toBeNull();
    });

    it("still marks a lock refreshed exactly one TTL ago", () => {
        expect(lockClassFor([held("sam", 2, 4)], "s1", 2, 4, 1_000 + LOCK_TTL_MS)).toBe(LOCK_CLASS);
    });

    it("paints nothing for a lock past its TTL", () => {
        expect(lockClassFor([held("sam", 2, 4)], "s1", 2, 4, 1_001 + LOCK_TTL_MS)).toBeNull();
    });

    it("marks each cell when two peers hold two different cells", () => {
        const locks = [held("sam", 2, 4), held("kim", 5, 9)];
        expect(lockClassFor(locks, "s1", 2, 4, 1_000)).toBe(LOCK_CLASS);
        expect(lockClassFor(locks, "s1", 5, 9, 1_000)).toBe(LOCK_CLASS);
        expect(lockClassFor(locks, "s1", 5, 4, 1_000)).toBeNull();
    });
});

describe("lockLabel", () => {
    it("names the holder", () => {
        expect(lockLabel([held("sam", 2, 4)], "s1", 2, 4, 1_000, nameOf)).toBe("Sam");
    });

    it("names each holder when two peers hold two different cells", () => {
        const locks = [held("sam", 2, 4), held("kim", 5, 9)];
        expect(lockLabel(locks, "s1", 2, 4, 1_000, nameOf)).toBe("Sam");
        expect(lockLabel(locks, "s1", 5, 9, 1_000, nameOf)).toBe("Kim");
    });

    it("falls back to whatever the resolver returns for an unknown peer", () => {
        expect(lockLabel([held("zed", 0, 0)], "s1", 0, 0, 1_000, nameOf)).toBe("zed");
    });

    it("names nobody on a free cell", () => {
        expect(lockLabel([held("sam", 2, 4)], "s1", 2, 5, 1_000, nameOf)).toBeNull();
    });

    it("names nobody on another sheet", () => {
        expect(lockLabel([held("sam", 2, 4)], "s2", 2, 4, 1_000, nameOf)).toBeNull();
    });

    it("names nobody once the lock has expired", () => {
        expect(lockLabel([held("sam", 2, 4)], "s1", 2, 4, 1_001 + LOCK_TTL_MS, nameOf)).toBeNull();
    });
});
