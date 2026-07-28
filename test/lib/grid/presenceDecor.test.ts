import { describe, expect, it } from "vitest";

import { PRESENCE_TTL_MS, type Presence } from "@/lib/collab/presence";
import { lockLabel, peerInitial, presenceOn } from "@/lib/grid/presenceDecor";

const held = (
    endpointId: string,
    col: number,
    row: number,
    heldAt = 1_000,
    sheetId = "s1",
    editing = true,
): Presence => ({
    endpointId,
    sheetId,
    col,
    row,
    heldAt,
    editing,
});

const NAMES: Record<string, string> = { sam: "Sam", kim: "Kim" };
const nameOf = (endpointId: string) => NAMES[endpointId] ?? endpointId;

describe("presenceOn", () => {
    it("finds the peer on a cell", () => {
        expect(presenceOn([held("sam", 2, 4)], "s1", 2, 4, 1_000)!.endpointId).toBe("sam");
    });

    it("finds a peer whose cursor is only resting there", () => {
        const parked = [held("sam", 2, 4, 1_000, "s1", false)];
        expect(presenceOn(parked, "s1", 2, 4, 1_000)!.editing).toBe(false);
    });

    it("leaves a neighbouring cell empty", () => {
        const list = [held("sam", 2, 4)];
        expect(presenceOn(list, "s1", 2, 5, 1_000)).toBeNull();
        expect(presenceOn(list, "s1", 3, 4, 1_000)).toBeNull();
    });

    it("leaves the same coordinates on another sheet empty", () => {
        expect(presenceOn([held("sam", 2, 4)], "s2", 2, 4, 1_000)).toBeNull();
    });

    it("finds nobody at all when no peer is anywhere", () => {
        expect(presenceOn([], "s1", 0, 0, 1_000)).toBeNull();
    });

    it("still finds a peer refreshed exactly one TTL ago", () => {
        expect(presenceOn([held("sam", 2, 4)], "s1", 2, 4, 1_000 + PRESENCE_TTL_MS)).not.toBeNull();
    });

    it("finds nobody past the TTL", () => {
        expect(presenceOn([held("sam", 2, 4)], "s1", 2, 4, 1_001 + PRESENCE_TTL_MS)).toBeNull();
    });

    it("finds each peer when two are on two different cells", () => {
        const list = [held("sam", 2, 4), held("kim", 5, 9)];
        expect(presenceOn(list, "s1", 2, 4, 1_000)!.endpointId).toBe("sam");
        expect(presenceOn(list, "s1", 5, 9, 1_000)!.endpointId).toBe("kim");
        expect(presenceOn(list, "s1", 5, 4, 1_000)).toBeNull();
    });
});

describe("peerInitial", () => {
    it("is the first letter of the name, in upper case", () => {
        expect(peerInitial("Sam")).toBe("S");
        expect(peerInitial("kim")).toBe("K");
    });

    it("skips leading punctuation and space", () => {
        expect(peerInitial("  rae")).toBe("R");
        expect(peerInitial('"Coach"')).toBe("C");
    });

    it("takes a digit, because a short EndpointId is what a nameless peer wears", () => {
        expect(peerInitial("3f9a2b1c")).toBe("3");
    });

    it("still leaves a mark for a name with no letter or digit in it", () => {
        expect(peerInitial("...")).toBe("*");
        expect(peerInitial("")).toBe("*");
    });
});

describe("lockLabel", () => {
    it("names the holder", () => {
        expect(lockLabel([held("sam", 2, 4)], "s1", 2, 4, 1_000, nameOf)).toBe("Sam");
    });

    it("names each holder when two peers hold two different cells", () => {
        const list = [held("sam", 2, 4), held("kim", 5, 9)];
        expect(lockLabel(list, "s1", 2, 4, 1_000, nameOf)).toBe("Sam");
        expect(lockLabel(list, "s1", 5, 9, 1_000, nameOf)).toBe("Kim");
    });

    it("names nobody for a cursor merely resting on the cell", () => {
        // The refusal hint answers a refused keystroke, and a resting cursor
        // refuses nothing.
        const parked = [held("sam", 2, 4, 1_000, "s1", false)];
        expect(lockLabel(parked, "s1", 2, 4, 1_000, nameOf)).toBeNull();
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

    it("names nobody once the claim has expired", () => {
        expect(
            lockLabel([held("sam", 2, 4)], "s1", 2, 4, 1_001 + PRESENCE_TTL_MS, nameOf),
        ).toBeNull();
    });
});
