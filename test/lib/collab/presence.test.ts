import { describe, expect, it } from "vitest";

import {
    claim,
    expire,
    HEARTBEAT_MS,
    lockAt,
    presenceAt,
    PRESENCE_TTL_MS,
    releaseCell,
    releasePeer,
    type Presence,
} from "@/lib/collab/presence";
import { modelCol } from "@/lib/grid/colSpace";

const at = (
    endpointId: string,
    row: number,
    heldAt = 0,
    sheetId = "s1",
    editing = true,
    readOnly = false,
): Presence => ({
    endpointId,
    sheetId,
    col: modelCol(0),
    row,
    heldAt,
    editing,
    readOnly,
});

describe("claim", () => {
    it("holds one cell for a peer", () => {
        expect(claim([], at("sam", 2))).toEqual([at("sam", 2)]);
    });

    it("replaces that peer's previous cell, because a cursor is on one", () => {
        const held = claim(claim([], at("sam", 2)), at("sam", 5));
        expect(held).toEqual([at("sam", 5)]);
    });

    it("replaces it whichever way round the two states arrive", () => {
        const parked = claim(claim([], at("sam", 2)), at("sam", 2, 0, "s1", false));
        expect(parked).toEqual([at("sam", 2, 0, "s1", false)]);
        const editing = claim(parked, at("sam", 2));
        expect(editing).toEqual([at("sam", 2)]);
    });

    it("lets two peers hold two cells at once", () => {
        const held = claim(claim([], at("sam", 2)), at("kim", 5));
        expect(held.map((p) => p.endpointId).sort()).toEqual(["kim", "sam"]);
    });
});

describe("release", () => {
    it("drops the one cell a peer is on when it says it is on none", () => {
        expect(releaseCell(claim([], at("sam", 2)), "sam")).toEqual([]);
    });

    it("drops everything a peer held when the connection goes", () => {
        const held = claim(claim([], at("sam", 2)), at("kim", 5));
        expect(releasePeer(held, "sam").map((p) => p.endpointId)).toEqual(["kim"]);
    });

    it("is a no-op for a peer that holds nothing", () => {
        expect(releasePeer([], "nobody")).toEqual([]);
        expect(releaseCell([], "nobody")).toEqual([]);
    });
});

describe("expire", () => {
    it("keeps an entry a heartbeat is still refreshing", () => {
        const held = claim([], at("sam", 2, 1_000));
        expect(expire(held, 1_000 + HEARTBEAT_MS, PRESENCE_TTL_MS)).toHaveLength(1);
    });

    it("drops an entry nothing has refreshed inside the window", () => {
        const held = claim([], at("sam", 2, 1_000));
        expect(expire(held, 1_000 + PRESENCE_TTL_MS + 1, PRESENCE_TTL_MS)).toEqual([]);
    });

    it("expires a resting cursor on the same clock as an open editor", () => {
        const parked = claim([], at("sam", 2, 1_000, "s1", false));
        expect(expire(parked, 1_000 + PRESENCE_TTL_MS + 1, PRESENCE_TTL_MS)).toEqual([]);
    });

    it("beats the TTL comfortably, so a live peer never flickers", () => {
        expect(HEARTBEAT_MS * 2).toBeLessThan(PRESENCE_TTL_MS);
    });

    it("leaves an unreachable peer holding nothing at all", () => {
        // A frozen process on a live connection is the only case the timer is
        // for; a dropped connection releases instantly through releasePeer.
        const held = claim([], at("gone", 0, 0));
        expect(expire(held, PRESENCE_TTL_MS + 1, PRESENCE_TTL_MS)).toEqual([]);
        expect(lockAt(expire(held, PRESENCE_TTL_MS + 1, PRESENCE_TTL_MS), "s1", 0, 0)).toBeNull();
    });
});

describe("presenceAt", () => {
    it("finds the peer on a cell whether or not they are editing it", () => {
        expect(presenceAt(claim([], at("sam", 2)), "s1", 0, 2)!.endpointId).toBe("sam");
        expect(presenceAt(claim([], at("sam", 2, 0, "s1", false)), "s1", 0, 2)!.editing).toBe(
            false,
        );
    });

    it("does not confuse the same coordinates on another sheet", () => {
        expect(presenceAt(claim([], at("sam", 2, 0, "s1")), "s2", 0, 2)).toBeNull();
    });

    it("reports nothing for a cell nobody is on", () => {
        expect(presenceAt(claim([], at("sam", 2)), "s1", 0, 3)).toBeNull();
    });
});

describe("lockAt", () => {
    it("finds the peer holding a cell", () => {
        expect(lockAt(claim([], at("sam", 2)), "s1", 0, 2)!.endpointId).toBe("sam");
    });

    it("holds nothing for a cursor merely resting on the cell", () => {
        // The whole point of the two states: a partner reading over your
        // shoulder must never refuse your keystroke.
        expect(lockAt(claim([], at("sam", 2, 0, "s1", false)), "s1", 0, 2)).toBeNull();
    });

    it("does not confuse the same coordinates on another sheet", () => {
        const held = claim([], at("sam", 2, 0, "s1"));
        expect(lockAt(held, "s2", 0, 2)).toBeNull();
    });

    it("reports nothing for a free cell", () => {
        expect(lockAt(claim([], at("sam", 2)), "s1", 0, 3)).toBeNull();
    });
});
