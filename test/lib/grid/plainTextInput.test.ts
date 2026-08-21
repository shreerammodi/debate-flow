import { expect, it } from "vitest";

import { disableTextAssistance, seedAppend } from "@/lib/grid/plainTextInput";

it("disables browser text assistance on the flow editor", () => {
    const input = document.createElement("textarea");

    disableTextAssistance(input);

    expect(input.getAttribute("autocorrect")).toBe("off");
    expect(input.getAttribute("autocapitalize")).toBe("off");
    expect(input.getAttribute("autocomplete")).toBe("off");
    expect(input.spellcheck).toBe(false);
});

it("seeds the editor with the cell's text and a caret past its end", () => {
    const input = document.createElement("textarea");

    seedAppend(input, "perm");

    expect(input.value).toBe("perm");
    expect(input.selectionStart).toBe(4);
    expect(input.selectionEnd).toBe(4);
});
