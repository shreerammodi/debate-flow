import { describe, expect, it } from "vitest";

import { hashText, sheetDigest } from "@/lib/collab/hash";

describe("hashText", () => {
    it("is stable for one input and differs for another", () => {
        expect(hashText("perm do both")).toBe(hashText("perm do both"));
        expect(hashText("perm do both")).not.toBe(hashText("perm do bath"));
    });

    it("separates the empty string from a null-ish one", () => {
        expect(hashText("")).not.toBe(hashText("0"));
    });

    it("reads as lowercase hex", () => {
        expect(hashText("anything")).toMatch(/^[0-9a-f]{8}$/);
    });
});

describe("sheetDigest", () => {
    it("ignores raggedness, because the grid pads to a rectangle anyway", () => {
        const ragged = sheetDigest([["a"], ["b", "c"]], {});
        const padded = sheetDigest(
            [
                ["a", null],
                ["b", "c"],
            ],
            {},
        );
        expect(ragged).toBe(padded);
    });

    it("ignores meta key order", () => {
        const one = sheetDigest([["a", "b"]], { "0,0": { bold: true }, "0,1": { card: true } });
        const two = sheetDigest([["a", "b"]], { "0,1": { card: true }, "0,0": { bold: true } });
        expect(one).toBe(two);
    });

    it("ignores an empty meta entry, which the grid drops on its own", () => {
        expect(sheetDigest([["a"]], { "0,0": {} })).toBe(sheetDigest([["a"]], {}));
    });

    it("treats a trailing empty row as absent", () => {
        expect(sheetDigest([["a"], [null]], {})).toBe(sheetDigest([["a"]], {}));
    });

    it("notices a changed cell, a moved cell, and a changed decoration", () => {
        const base = sheetDigest([["a", "b"]], { "0,0": { bold: true } });
        expect(sheetDigest([["a", "c"]], { "0,0": { bold: true } })).not.toBe(base);
        expect(sheetDigest([["b", "a"]], { "0,0": { bold: true } })).not.toBe(base);
        expect(sheetDigest([["a", "b"]], { "0,0": { card: true } })).not.toBe(base);
        expect(sheetDigest([["a", "b"]], {})).not.toBe(base);
    });

    it("does not confuse a cell boundary with cell text", () => {
        expect(sheetDigest([["a", "b"]], {})).not.toBe(sheetDigest([["a\u0000b"]], {}));
    });
});
