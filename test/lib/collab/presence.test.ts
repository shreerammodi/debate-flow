import { describe, expect, it } from "vitest";

import {
    claim,
    expire,
    HEARTBEAT_MS,
    lockAt,
    LOCK_TTL_MS,
    releaseCell,
    releasePeer,
    type Lock,
} from "@/lib/collab/presence";

const at = (endpointId: string, row: number, heldAt = 0, sheetId = "s1"): Lock => ({
    endpointId,
    sheetId,
    col: 0,
    row,
    heldAt,
});

describe("claim", () => {
    it("holds one cell for a peer", () => {
        expect(claim([], at("sam", 2))).toEqual([at("sam", 2)]);
    });

    it("replaces that peer's previous cell, because an editor opens on one", () => {
        const held = claim(claim([], at("sam", 2)), at("sam", 5));
        expect(held).toEqual([at("sam", 5)]);
    });

    it("lets two peers hold two cells at once", () => {
        const held = claim(claim([], at("sam", 2)), at("kim", 5));
        expect(held.map((l) => l.endpointId).sort()).toEqual(["kim", "sam"]);
    });
});

describe("release", () => {
    it("drops the one cell a peer holds when its editor closes", () => {
        expect(releaseCell(claim([], at("sam", 2)), "sam")).toEqual([]);
    });

    it("drops every lock a peer held when the connection goes", () => {
        const held = claim(claim([], at("sam", 2)), at("kim", 5));
        expect(releasePeer(held, "sam").map((l) => l.endpointId)).toEqual(["kim"]);
    });

    it("is a no-op for a peer that holds nothing", () => {
        expect(releasePeer([], "nobody")).toEqual([]);
        expect(releaseCell([], "nobody")).toEqual([]);
    });
});

describe("expire", () => {
    it("keeps a lock a heartbeat is still refreshing", () => {
        const held = claim([], at("sam", 2, 1_000));
        expect(expire(held, 1_000 + HEARTBEAT_MS, LOCK_TTL_MS)).toHaveLength(1);
    });

    it("drops a lock nothing has refreshed inside the window", () => {
        const held = claim([], at("sam", 2, 1_000));
        expect(expire(held, 1_000 + LOCK_TTL_MS + 1, LOCK_TTL_MS)).toEqual([]);
    });

    it("beats the TTL comfortably, so a live peer never flickers", () => {
        expect(HEARTBEAT_MS * 2).toBeLessThan(LOCK_TTL_MS);
    });

    it("leaves an unreachable peer holding nothing at all", () => {
        // A frozen process on a live connection is the only case the timer is
        // for; a dropped connection releases instantly through releasePeer.
        const held = claim([], at("gone", 0, 0));
        expect(expire(held, LOCK_TTL_MS + 1, LOCK_TTL_MS)).toEqual([]);
        expect(lockAt(expire(held, LOCK_TTL_MS + 1, LOCK_TTL_MS), "s1", 0, 0)).toBeNull();
    });
});

describe("lockAt", () => {
    it("finds the peer holding a cell", () => {
        expect(lockAt(claim([], at("sam", 2)), "s1", 0, 2)!.endpointId).toBe("sam");
    });

    it("does not confuse the same coordinates on another sheet", () => {
        const held = claim([], at("sam", 2, 0, "s1"));
        expect(lockAt(held, "s2", 0, 2)).toBeNull();
    });

    it("reports nothing for a free cell", () => {
        expect(lockAt(claim([], at("sam", 2)), "s1", 0, 3)).toBeNull();
    });
});
