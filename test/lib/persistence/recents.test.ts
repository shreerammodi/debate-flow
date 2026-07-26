import { describe, expect, it } from "vitest";

import {
    dropRecent,
    parseRecents,
    promoteRecent,
    RECENTS_KEPT,
    serializeRecents,
    type RecentFlow,
} from "@/lib/persistence/recents";

const entry = (path: string, openedAt = 1): RecentFlow => ({ path, openedAt });

describe("promoteRecent", () => {
    it("puts the newest first", () => {
        const list = promoteRecent([entry("/a")], "/b", 5);
        expect(list.map((r) => r.path)).toEqual(["/b", "/a"]);
    });

    it("moves an existing path rather than duplicating it", () => {
        const list = promoteRecent([entry("/a"), entry("/b")], "/b", 9);
        expect(list.map((r) => r.path)).toEqual(["/b", "/a"]);
        expect(list[0].openedAt).toBe(9);
    });

    it("caps the list so it cannot grow without bound", () => {
        let list: RecentFlow[] = [];
        for (let i = 0; i < RECENTS_KEPT + 5; i++) list = promoteRecent(list, `/f${i}`, i);
        expect(list).toHaveLength(RECENTS_KEPT);
        expect(list[0].path).toBe(`/f${RECENTS_KEPT + 4}`);
    });
});

describe("dropRecent", () => {
    it("removes only the named path", () => {
        expect(dropRecent([entry("/a"), entry("/b")], "/a").map((r) => r.path)).toEqual(["/b"]);
    });
});

describe("parseRecents", () => {
    it("round-trips what serializeRecents wrote", () => {
        const list = [entry("/a", 2), entry("/b", 1)];
        expect(parseRecents(serializeRecents(list))).toEqual(list);
    });

    it("degrades to empty rather than blocking the start screen", () => {
        // The file is hand-editable and syncable, so every one of these is
        // reachable in the wild.
        expect(parseRecents(null)).toEqual([]);
        expect(parseRecents("")).toEqual([]);
        expect(parseRecents("{oh no")).toEqual([]);
        expect(parseRecents("[]")).toEqual([]);
        expect(parseRecents('{"flows":"nope"}')).toEqual([]);
    });

    it("skips unusable entries but keeps the good ones around them", () => {
        const text = JSON.stringify({
            version: 1,
            flows: [{ path: 7 }, null, { path: "" }, { path: "/good" }],
        });
        expect(parseRecents(text)).toEqual([{ path: "/good", openedAt: 0 }]);
    });

    it("defaults a missing or non-numeric timestamp", () => {
        const text = JSON.stringify({ flows: [{ path: "/a" }, { path: "/b", openedAt: "soon" }] });
        expect(parseRecents(text)).toEqual([
            { path: "/a", openedAt: 0 },
            { path: "/b", openedAt: 0 },
        ]);
    });

    it("collapses a duplicated path to its first appearance", () => {
        const text = JSON.stringify({
            flows: [
                { path: "/a", openedAt: 3 },
                { path: "/a", openedAt: 1 },
            ],
        });
        expect(parseRecents(text)).toEqual([{ path: "/a", openedAt: 3 }]);
    });
});
