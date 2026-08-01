import { expect, it } from "vitest";

import { disableTextAssistance } from "@/lib/grid/plainTextInput";

it("disables browser text assistance on the flow editor", () => {
    const input = document.createElement("textarea");

    disableTextAssistance(input);

    expect(input.getAttribute("autocorrect")).toBe("off");
    expect(input.getAttribute("autocapitalize")).toBe("off");
    expect(input.getAttribute("autocomplete")).toBe("off");
    expect(input.spellcheck).toBe(false);
});
