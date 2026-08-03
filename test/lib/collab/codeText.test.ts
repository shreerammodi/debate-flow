import { describe, expect, it } from "vitest";

import { bareCode, groupCode, looksLikeCode } from "@/lib/collab/codeText";

describe("groupCode", () => {
    it("splits eight characters down the middle", () => {
        expect(groupCode("K7QM3XPV")).toBe("K7QM-3XPV");
    });

    it("leaves a code that is already grouped alone", () => {
        expect(groupCode("K7QM-3XPV")).toBe("K7QM-3XPV");
    });

    it("shows a partial code as it is, rather than guessing at a group", () => {
        expect(groupCode("K7Q")).toBe("K7Q");
        expect(groupCode("")).toBe("");
    });
});

describe("bareCode", () => {
    it("hands back the eight characters the shell derives from", () => {
        expect(bareCode(" k7qm-3xpv ")).toBe("K7QM3XPV");
    });
});

describe("looksLikeCode", () => {
    it("takes a code however it is typed", () => {
        expect(looksLikeCode("k7qm3xpv")).toBe(true);
        expect(looksLikeCode("K7QM-3XPV")).toBe(true);
        expect(looksLikeCode(" k7qm 3xpv ")).toBe(true);
    });

    it("refuses the characters ebb's codes never use", () => {
        for (const bad of ["K7QM3XPI", "K7QM3XPL", "K7QM3XPO", "K7QM3XPU"]) {
            expect(looksLikeCode(bad)).toBe(false);
        }
    });

    it("refuses anything that is not eight characters", () => {
        expect(looksLikeCode("K7QM3XP")).toBe(false);
        expect(looksLikeCode("K7QM3XPVV")).toBe(false);
        expect(looksLikeCode("")).toBe(false);
    });
});
