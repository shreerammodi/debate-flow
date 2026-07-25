import { describe, expect, it } from "vitest";

import { parseFlowRequest, parseRevealRequest } from "@/lib/bridge/protocol";

describe("parseFlowRequest", () => {
    it("reads a well-formed send", () => {
        const req = parseFlowRequest({
            mode: "cell",
            docTitle: "AT - Cap K",
            items: [{ kind: "tag", text: "Perm solves", source: "cmsrc1.a", key: "doc-1|perm" }],
            space: 2,
        });
        expect(req).toEqual({
            mode: "cell",
            docTitle: "AT - Cap K",
            items: [{ kind: "tag", text: "Perm solves", token: "cmsrc1.a", key: "doc-1|perm" }],
            space: 2,
        });
    });

    it("defaults the mode to column and the kind to analytic", () => {
        const req = parseFlowRequest({ mode: "sideways", items: [{ text: "No link" }] });
        expect(req?.mode).toBe("column");
        expect(req?.items[0]).toEqual({ kind: "analytic", text: "No link", token: "", key: "" });
        expect(req?.docTitle).toBe("");
    });

    it("limits the empty-cell count, and reads a missing or bad one as none", () => {
        const space = (value: unknown) =>
            parseFlowRequest({ items: [{ text: "Kept" }], space: value })?.space;
        expect(space(99)).toBe(10);
        expect(space(-3)).toBe(0);
        expect(space(1.6)).toBe(2);
        expect(space("two")).toBe(0);
        expect(parseFlowRequest({ items: [{ text: "Kept" }] })?.space).toBe(0);
    });

    it("drops items with no text but keeps the rest", () => {
        const req = parseFlowRequest({ items: [{ text: "" }, "junk", { text: "Kept" }] });
        expect(req?.items.map((i) => i.text)).toEqual(["Kept"]);
    });

    it("rejects a body with no usable items", () => {
        expect(parseFlowRequest(null)).toBeNull();
        expect(parseFlowRequest({})).toBeNull();
        expect(parseFlowRequest({ items: "nope" })).toBeNull();
        expect(parseFlowRequest({ items: [] })).toBeNull();
        expect(parseFlowRequest({ items: [{ text: "" }] })).toBeNull();
    });
});

describe("parseRevealRequest", () => {
    it("keeps only the non-empty string keys", () => {
        expect(parseRevealRequest({ keys: ["a", "", 7, "b"], docTitle: "AT" })).toEqual({
            keys: ["a", "b"],
            docTitle: "AT",
        });
    });

    it("rejects a body with no usable keys", () => {
        expect(parseRevealRequest({})).toBeNull();
        expect(parseRevealRequest({ keys: [] })).toBeNull();
        expect(parseRevealRequest({ keys: [1, 2] })).toBeNull();
    });
});
