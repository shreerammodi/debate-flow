import { describe, expect, it } from "vitest";

import type { EventId } from "@/lib/format/events";
import { emptyScouting } from "@/lib/model/flow";
import type { Scouting } from "@/lib/model/types";
import {
    basename,
    dedupeFilename,
    dirname,
    displayPath,
    joinPath,
    stem,
    suggestFilename,
    withEbbExt,
} from "@/lib/persistence/flowPaths";

function scouted(over: Partial<Scouting> = {}): Scouting {
    return {
        ...emptyScouting(),
        affSchool: "Westwood",
        negSchool: "Harvard",
        aff: { first: { first: "Ada", last: "Gray" }, second: { first: "Max", last: "Moss" } },
        neg: { first: { first: "Ben", last: "Stone" }, second: { first: "Cy", last: "Byrd" } },
        ...over,
    };
}

describe("path splitting", () => {
    it("handles both separators, because Windows paths arrive verbatim", () => {
        expect(basename("/home/a/round.ebb")).toBe("round.ebb");
        expect(basename("C:\\Users\\a\\round.ebb")).toBe("round.ebb");
        expect(dirname("/home/a/round.ebb")).toBe("/home/a");
        expect(dirname("C:\\Users\\a\\round.ebb")).toBe("C:\\Users\\a");
    });

    it("treats a bare name as having no directory", () => {
        expect(basename("round.ebb")).toBe("round.ebb");
        expect(dirname("round.ebb")).toBe("");
    });

    it("strips only the .ebb extension, case-insensitively", () => {
        expect(stem("/a/round.ebb")).toBe("round");
        expect(stem("/a/round.EBB")).toBe("round");
        expect(stem("/a/round.json")).toBe("round.json");
    });

    it("never doubles the extension", () => {
        expect(withEbbExt("round")).toBe("round.ebb");
        expect(withEbbExt("round.ebb")).toBe("round.ebb");
    });

    it("joins with the separator the directory already uses", () => {
        expect(joinPath("/home/a", "r.ebb")).toBe("/home/a/r.ebb");
        expect(joinPath("/home/a/", "r.ebb")).toBe("/home/a/r.ebb");
        expect(joinPath("C:\\Users\\a", "r.ebb")).toBe("C:\\Users\\a\\r.ebb");
        expect(joinPath("", "r.ebb")).toBe("r.ebb");
    });
});

describe("displayPath", () => {
    it("collapses the home directory", () => {
        expect(displayPath("/home/test/Documents/r.ebb", "/home/test")).toBe("~/Documents/r.ebb");
        expect(displayPath("/home/test", "/home/test")).toBe("~");
    });

    it("leaves a path outside home alone", () => {
        expect(displayPath("/srv/r.ebb", "/home/test")).toBe("/srv/r.ebb");
    });

    it("does not collapse a sibling whose name merely starts with home", () => {
        expect(displayPath("/home/test2/r.ebb", "/home/test")).toBe("/home/test2/r.ebb");
    });
});

describe("suggestFilename", () => {
    it("names a scouted round after the matchup", () => {
        const name = suggestFilename({
            event: "policy",
            createdAt: Date.UTC(2026, 6, 25, 12),
            scouting: scouted({ tournament: "Berkeley", round: "Round 3" }),
        });
        expect(name).toBe("berkeley-round-3-westwood-gm-vs-harvard-bs.ebb");
    });

    it("falls back to event and date, which is what a brand-new flow has", () => {
        const created = new Date(2026, 6, 25, 20, 30).getTime();
        expect(
            suggestFilename({ event: "pf", createdAt: created, scouting: emptyScouting() }),
        ).toBe("pf-2026-07-25.ebb");
    });

    it("uses the local date, so an evening round is not filed tomorrow", () => {
        // 23:30 local on the 25th is the 26th in UTC.
        const created = new Date(2026, 6, 25, 23, 30).getTime();
        expect(
            suggestFilename({ event: "policy", createdAt: created, scouting: emptyScouting() }),
        ).toContain("2026-07-25");
    });

    it("keeps a partially scouted round useful", () => {
        const name = suggestFilename({
            event: "ld",
            createdAt: 0,
            scouting: { ...emptyScouting(), tournament: "Glenbrooks" },
        });
        expect(name).toBe("glenbrooks.ebb");
    });

    // A guest's round is the host's document flattened, so `event` is whatever
    // the host put on the wire. The name is then joined onto the flows
    // directory, where a separator or a `..` segment would leave it.
    it("cannot be steered out of the flows directory by a hostile event", () => {
        const created = new Date(2026, 6, 25, 12).getTime();
        for (const event of [
            "../../../../tmp/evil",
            "/Users/victim/Library/Application Support/evil",
            "..\\..\\evil",
            "..",
            "/",
            "",
        ]) {
            const name = suggestFilename({
                event: event as EventId,
                createdAt: created,
                scouting: emptyScouting(),
            });
            expect(name, event).toMatch(/^[a-z0-9-]+\.ebb$/);
            expect(name, event).not.toContain("..");
        }

        // An event with nothing sluggable left takes the ordinary default, so
        // the guest lands on a normally named flow rather than a refusal.
        for (const event of ["..", "/", "", "///"]) {
            expect(
                suggestFilename({
                    event: event as EventId,
                    createdAt: created,
                    scouting: emptyScouting(),
                }),
                event,
            ).toBe("policy-2026-07-25.ebb");
        }
    });

    it("bounds the fallback name, which the host also chooses", () => {
        const name = suggestFilename({
            event: "z".repeat(500) as EventId,
            createdAt: 0,
            scouting: emptyScouting(),
        });
        expect(name).toBe(`${"z".repeat(72)}.ebb`);
    });
});

describe("dedupeFilename", () => {
    it("leaves a free name alone", () => {
        expect(dedupeFilename("r.ebb", new Set())).toBe("r.ebb");
    });

    it("walks past every taken variant", () => {
        const taken = new Set(["r.ebb", "r-2.ebb", "r-3.ebb"]);
        expect(dedupeFilename("r.ebb", taken)).toBe("r-4.ebb");
    });
});
