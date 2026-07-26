import { describe, expect, it } from "vitest";

import { errorMessage } from "@/lib/errorMessage";

describe("errorMessage", () => {
    it("keeps the string a Tauri command rejected with", () => {
        // invoke() rejects with the plain Err string, not an Error. Treating
        // that as unrecognized is what replaced every Rust diagnostic with a
        // generic fallback.
        expect(errorMessage("Could not create /Users/a/Documents/ebb: denied", "fallback")).toBe(
            "Could not create /Users/a/Documents/ebb: denied",
        );
    });

    it("keeps a thrown Error's message", () => {
        expect(errorMessage(new Error("Not a flow file"), "fallback")).toBe("Not a flow file");
    });

    it("falls back for anything without usable text", () => {
        expect(errorMessage(undefined, "fallback")).toBe("fallback");
        expect(errorMessage(null, "fallback")).toBe("fallback");
        expect(errorMessage({ code: 7 }, "fallback")).toBe("fallback");
        expect(errorMessage("", "fallback")).toBe("fallback");
        expect(errorMessage("   ", "fallback")).toBe("fallback");
        expect(errorMessage(new Error(""), "fallback")).toBe("fallback");
    });
});
