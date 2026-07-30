/**
 * The sanitizer a peer's RFD markdown goes through.
 *
 * An RFD section arrives inside a peer's document and is rendered as HTML, so
 * this is the one place between a partner's bytes and the webview. Its Rust
 * half - `navigable_in` plus `on_new_window` in `src-tauri/src/windows.rs` -
 * exists only in the desktop shell, and `FORBID_ATTR: ["href"]` is what covers
 * the schemes that guard deliberately leaves open and the browser build has no
 * guard at all for.
 */

import { describe, expect, it } from "vitest";

import { renderRfdHtml } from "@/lib/rfd/markdown";

describe("renderRfdHtml", () => {
    it("renders the notes a judge actually writes", () => {
        const html = renderRfdHtml("Aff wins on **turns case**.\n\n- clean impact\n");
        expect(html).toContain("<strong>turns case</strong>");
        expect(html).toContain("<li>clean impact</li>");
    });

    it("strips the href off a link, so a note cannot navigate the app away", () => {
        const html = renderRfdHtml("[click](javascript:alert(1))");
        expect(html).not.toContain("javascript:");
        expect(html).not.toContain("href");
        // The text survives; only the navigation is gone.
        expect(html).toContain("click");
    });

    it("strips the href off an ordinary link too, and every scheme with it", () => {
        // Nothing in an RFD navigates. DOMPurify's own URI policy passes
        // mailto:, tel: and cid:, which the shell forwards to the OS, and
        // `navigable_in`'s final arm is deliberately open. This is what closes
        // both, and it is the one assertion FORBID_ATTR alone defends: the
        // javascript: case above is refused by the default policy as well.
        expect(renderRfdHtml('<a href="mailto:judge@example.com">mail</a>')).not.toContain(
            "mailto:",
        );
        const http = renderRfdHtml("[their evidence](https://evil.example/doc)");
        expect(http).not.toContain("href");
        expect(http).not.toContain("evil.example");
        expect(http).toContain("their evidence");
    });

    it("drops an event handler rather than rendering it", () => {
        const html = renderRfdHtml('<img src=x onerror="alert(1)">');
        expect(html).not.toContain("onerror");
        expect(html).not.toContain("alert");
    });

    it("drops an embedded frame outright", () => {
        const html = renderRfdHtml('<iframe src="https://evil.example"></iframe>');
        expect(html).not.toContain("iframe");
        expect(html).not.toContain("evil.example");
    });

    it("drops target, which is the half the navigation guard is paired with", () => {
        // `on_new_window` closes the window.open route in the shell; target is
        // the markup route to the same thing, and the two must not disagree.
        const html = renderRfdHtml('<a href="https://evil.example" target="_blank">x</a>');
        expect(html).not.toContain("target");
    });

    it("drops a script even when the markdown hands it over verbatim", () => {
        const html = renderRfdHtml('<script>alert("rfd")</script>\n\nreal notes\n');
        expect(html).not.toContain("<script");
        expect(html).toContain("real notes");
    });
});
