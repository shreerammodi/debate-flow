import { afterEach, describe, expect, it, vi } from "vitest";

import type { Presence } from "@/lib/collab/presence";
import { modelCol } from "@/lib/grid/colSpace";
import {
    claimCell,
    claimCursor,
    getPresences,
    onPresenceChanged,
    setClaimHandler,
    setCursorHandler,
    setPresences,
    type HeldCell,
} from "@/lib/grid/presenceBridge";

afterEach(() => {
    setPresences([]);
    setClaimHandler(null);
    setCursorHandler(null);
});

const on = (endpointId: string, editing = true, readOnly = false): Presence => ({
    endpointId,
    sheetId: "sheet_1",
    col: modelCol(0),
    row: 3,
    heldAt: 1_000,
    editing,
    readOnly,
});

describe("presence arriving from a session", () => {
    it("is held and announced to every mounted pane", () => {
        const one = vi.fn();
        const two = vi.fn();
        const drop = onPresenceChanged(one);
        onPresenceChanged(two);

        setPresences([on("sam")]);

        expect(getPresences()).toEqual([on("sam")]);
        expect(one).toHaveBeenCalledTimes(1);
        expect(two).toHaveBeenCalledTimes(1);

        drop();
        setPresences([]);
        expect(one).toHaveBeenCalledTimes(1);
        expect(two).toHaveBeenCalledTimes(2);
    });

    it("repaints for a peer that only moved its cursor", () => {
        const paint = vi.fn();
        onPresenceChanged(paint);
        setPresences([on("sam", false)]);
        expect(getPresences()[0].editing).toBe(false);
        expect(paint).toHaveBeenCalledTimes(1);
    });

    it("hands back one array for every empty table, so a clear is not a change", () => {
        setPresences([]);
        const first = getPresences();
        setPresences([]);
        expect(getPresences()).toBe(first);
    });
});

describe("the cell this side is editing", () => {
    it("reaches whatever session is listening", () => {
        const claims: HeldCell[] = [];
        setClaimHandler((cell) => claims.push(cell));

        claimCell({ sheetId: "sheet_1", col: modelCol(2), row: 7 });
        claimCell(null);

        expect(claims).toEqual([{ sheetId: "sheet_1", col: 2, row: 7 }, null]);
    });

    // A debater flowing alone announces nothing to anybody, and the grid does
    // not know whether anyone is listening.
    it("is a no-op with no session", () => {
        expect(() => claimCell({ sheetId: "sheet_1", col: modelCol(0), row: 0 })).not.toThrow();
    });

    it("stops reaching a session that has ended", () => {
        const claims: HeldCell[] = [];
        setClaimHandler((cell) => claims.push(cell));
        claimCell({ sheetId: "sheet_1", col: modelCol(0), row: 0 });
        setClaimHandler(null);
        claimCell({ sheetId: "sheet_1", col: modelCol(1), row: 1 });
        expect(claims).toHaveLength(1);
    });
});

describe("the cell this side's cursor is on", () => {
    it("travels its own route, so an editor claim is not implied by a selection", () => {
        const claims: HeldCell[] = [];
        const cursors: HeldCell[] = [];
        setClaimHandler((cell) => claims.push(cell));
        setCursorHandler((cell) => cursors.push(cell));

        claimCursor({ sheetId: "sheet_1", col: modelCol(1), row: 4 });
        claimCursor(null);

        expect(cursors).toEqual([{ sheetId: "sheet_1", col: 1, row: 4 }, null]);
        expect(claims).toEqual([]);
    });

    it("is a no-op with no session", () => {
        expect(() => claimCursor({ sheetId: "sheet_1", col: modelCol(0), row: 0 })).not.toThrow();
    });

    it("stops reaching a session that has ended", () => {
        const cursors: HeldCell[] = [];
        setCursorHandler((cell) => cursors.push(cell));
        claimCursor({ sheetId: "sheet_1", col: modelCol(0), row: 0 });
        setCursorHandler(null);
        claimCursor({ sheetId: "sheet_1", col: modelCol(1), row: 1 });
        expect(cursors).toHaveLength(1);
    });
});
