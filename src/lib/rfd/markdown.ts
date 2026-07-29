/**
 * Markdown to display HTML for one RFD section.
 *
 * Every author's notes go through this on their own, never concatenated first.
 * An RFD is a trust boundary even when it is only an imported flow's, and a
 * peer's is a stricter one: sanitizing per section means an unterminated
 * construct in one author's markdown closes with that author's fragment
 * instead of swallowing the next author's section.
 */

import DOMPurify from "dompurify";
import { marked } from "marked";

export function renderRfdHtml(text: string): string {
    // Nothing in an RFD navigates. The app's own outbound links go through the
    // scoped opener; a peer's note is markdown from a trust boundary, and a
    // top-level navigation is the one thing the CSP does not restrain, so a
    // click on their link would replace the flowing app with their page.
    return DOMPurify.sanitize(marked.parse(text) as string, { FORBID_ATTR: ["href"] });
}
