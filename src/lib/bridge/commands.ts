/**
 * The two commands that reach out to CardMirror: jump to a cell's source
 * document position, and send the selection into the open document.
 *
 * Both are fire-and-forget from the command layer's point of view. They never
 * throw and never leave the user guessing: every outcome, including a
 * CardMirror that is closed or a cell with no provenance, ends in a toast
 * that names the next move.
 */

import { toast } from "sonner";

import { getActiveHot } from "@/lib/grid/hotInstance";
import type { CellSource } from "@/lib/model/flow";
import { useFlowStore } from "@/lib/store/useFlowStore";

import type { BridgeCall, BridgeFailure, CardMirrorReply } from "./cardmirror";
import { cardmirrorInsert, cardmirrorJump } from "./cardmirror";

const TRANSPORT_MESSAGE: Record<BridgeFailure, string> = {
    "not-registered": "CardMirror has never run on this machine.",
    "not-running": "CardMirror is not running.",
    timeout: "CardMirror did not answer.",
    "bad-response": "CardMirror sent something ebb could not read.",
    unsupported: "This works in the ebb desktop app.",
};

const JUMP_MESSAGE: Record<string, string> = {
    "not-found": "That card is no longer in the document.",
    "bad-request": "CardMirror could not read this cell's source.",
};

const INSERT_MESSAGE: Record<string, string> = {
    "no-target-doc": "Open a document in CardMirror first.",
    "doc-readonly": "That CardMirror document is in read mode.",
    "bad-request": "CardMirror would not take that text.",
};

/** The provenance on the focused cell, or null when it was typed here. */
function selectedSource(): CellSource | null {
    const hot = getActiveHot();
    const selection = hot?.getSelectedLast();
    if (!hot || !selection) return null;
    // Handsontable types cell meta as an open bag, so the read is asserted the
    // same way the rest of the grid layer asserts className.
    return (hot.getCellMeta(selection[0], selection[1]).source as CellSource | undefined) ?? null;
}

/**
 * Every non-empty selected cell, row-major within each selected range, one
 * per line. CardMirror's insert builds one block per line, so a single
 * newline is the paragraph break; a blank line between cells would leave an
 * empty paragraph behind in the document.
 */
function selectedText(): string {
    const hot = getActiveHot();
    const ranges = hot?.getSelectedRange();
    if (!hot || !ranges) return "";
    const parts: string[] = [];
    for (const range of ranges) {
        const tl = range.getTopLeftCorner();
        const br = range.getBottomRightCorner();
        for (let r = tl.row ?? 0; r <= (br.row ?? -1); r++) {
            for (let c = tl.col ?? 0; c <= (br.col ?? -1); c++) {
                const value = hot.getDataAtCell(r, c);
                if (typeof value === "string" && value.trim()) parts.push(value.trim());
            }
        }
    }
    return parts.join("\n");
}

/** Null when the call and CardMirror both succeeded, else what to tell the user. */
function failureMessage(
    call: BridgeCall<CardMirrorReply>,
    errors: Record<string, string>,
    fallback: string,
): string | null {
    if (!call.ok) return TRANSPORT_MESSAGE[call.error];
    if (call.value.ok) return null;
    const error = call.value.error ?? "";
    if (error === "doc-not-open") {
        const title = call.value.docTitle;
        return title ? `Open "${title}" in CardMirror first.` : "Open the document in CardMirror.";
    }
    return errors[error] ?? fallback;
}

export async function runJumpToSource(): Promise<void> {
    const source = selectedSource();
    if (!source) {
        toast("This cell did not come from CardMirror.");
        return;
    }
    const message = failureMessage(
        await cardmirrorJump(source.token),
        JUMP_MESSAGE,
        "CardMirror could not open this cell's source.",
    );
    if (message) toast(message);
}

export async function runSendToDoc(): Promise<void> {
    const text = selectedText();
    if (!text) {
        toast("Select a cell with text to send.");
        return;
    }
    const call = await cardmirrorInsert(text, useFlowStore.getState().cardmirrorTextType);
    const message = failureMessage(call, INSERT_MESSAGE, "CardMirror could not take that text.");
    if (message) {
        toast(message);
        return;
    }
    const title = call.ok ? call.value.docTitle : undefined;
    toast(title ? `Sent to "${title}".` : "Sent to CardMirror.");
}
