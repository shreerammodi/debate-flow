import { afterEach, describe, expect, it, vi } from "vitest";

import type { Lock } from "@/lib/collab/presence";
import {
    claimCell,
    getLocks,
    onLocksChanged,
    setClaimHandler,
    setLocks,
    type HeldCell,
} from "@/lib/grid/lockBridge";

afterEach(() => {
    setLocks([]);
    setClaimHandler(null);
});

const lock = (endpointId: string): Lock => ({
    endpointId,
    sheetId: "sheet_1",
    col: 0,
    row: 3,
    heldAt: 1_000,
});

describe("locks arriving from a session", () => {
    it("are held and announced to every mounted pane", () => {
        const one = vi.fn();
        const two = vi.fn();
        const drop = onLocksChanged(one);
        onLocksChanged(two);

        setLocks([lock("sam")]);

        expect(getLocks()).toEqual([lock("sam")]);
        expect(one).toHaveBeenCalledTimes(1);
        expect(two).toHaveBeenCalledTimes(1);

        drop();
        setLocks([]);
        expect(one).toHaveBeenCalledTimes(1);
        expect(two).toHaveBeenCalledTimes(2);
    });

    it("hand back one array for every empty table, so a clear is not a change", () => {
        setLocks([]);
        const first = getLocks();
        setLocks([]);
        expect(getLocks()).toBe(first);
    });
});

describe("the cell this side is editing", () => {
    it("reaches whatever session is listening", () => {
        const claims: HeldCell[] = [];
        setClaimHandler((cell) => claims.push(cell));

        claimCell({ sheetId: "sheet_1", col: 2, row: 7 });
        claimCell(null);

        expect(claims).toEqual([{ sheetId: "sheet_1", col: 2, row: 7 }, null]);
    });

    // A debater flowing alone announces nothing to anybody, and the grid does
    // not know whether anyone is listening.
    it("is a no-op with no session", () => {
        expect(() => claimCell({ sheetId: "sheet_1", col: 0, row: 0 })).not.toThrow();
    });

    it("stops reaching a session that has ended", () => {
        const claims: HeldCell[] = [];
        setClaimHandler((cell) => claims.push(cell));
        claimCell({ sheetId: "sheet_1", col: 0, row: 0 });
        setClaimHandler(null);
        claimCell({ sheetId: "sheet_1", col: 1, row: 1 });
        expect(claims).toHaveLength(1);
    });
});
