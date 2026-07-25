import { invoke } from "@tauri-apps/api/core";
import { toast } from "sonner";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { runJumpToSource, runSendToDoc } from "@/lib/bridge/commands";
import { setActiveHot } from "@/lib/grid/hotInstance";
import type { CellSource } from "@/lib/model/flow";
import { useFlowStore } from "@/lib/store/useFlowStore";

vi.mock("@tauri-apps/api/core", () => ({ invoke: vi.fn() }));
vi.mock("sonner", () => ({ toast: vi.fn() }));

const invoked = vi.mocked(invoke);
const toasted = vi.mocked(toast);

const SOURCE: CellSource = {
    app: "cardmirror",
    token: "cmsrc1.a",
    key: "doc-1|perm solves",
    title: "AT - Cap K",
};

/** A one-cell grid whose text and provenance the test controls. */
function makeGrid(text: string | null, source?: CellSource) {
    return {
        getSelectedLast: () => [0, 0],
        getSelectedRange: () => [
            {
                getTopLeftCorner: () => ({ row: 0, col: 0 }),
                getBottomRightCorner: () => ({ row: 0, col: 0 }),
            },
        ],
        getCellMeta: () => ({ source }),
        getDataAtCell: () => text,
    };
}

const lastToast = () => toasted.mock.calls.at(-1)?.[0];

beforeEach(() => {
    invoked.mockReset();
    toasted.mockReset();
    setActiveHot(null, null);
    (window as unknown as Record<string, unknown>).__TAURI_INTERNALS__ = {};
    useFlowStore.setState({ cardmirrorTextType: "analytic", cardmirrorEnabled: true });
});

afterEach(() => {
    delete (window as unknown as Record<string, unknown>).__TAURI_INTERNALS__;
});

describe("jump to source", () => {
    it("hands the cell's token to the host and stays quiet on success", async () => {
        setActiveHot(makeGrid("Perm solves", SOURCE) as never, vi.fn());
        invoked.mockResolvedValue({ ok: true });

        await runJumpToSource();

        expect(invoked).toHaveBeenCalledWith("cardmirror_jump", { source: "cmsrc1.a" });
        expect(toasted).not.toHaveBeenCalled();
    });

    it("says so on a cell that was typed here", async () => {
        setActiveHot(makeGrid("typed") as never, vi.fn());
        await runJumpToSource();
        expect(invoked).not.toHaveBeenCalled();
        expect(lastToast()).toBe("This cell did not come from CardMirror.");
    });

    it("names the document to open when it is closed", async () => {
        setActiveHot(makeGrid("Perm solves", SOURCE) as never, vi.fn());
        invoked.mockResolvedValue({ ok: false, error: "doc-not-open", docTitle: "AT - Cap K" });

        await runJumpToSource();
        expect(lastToast()).toBe('Open "AT - Cap K" in CardMirror first.');
    });

    it("explains a card that is gone and a source it cannot read", async () => {
        setActiveHot(makeGrid("Perm solves", SOURCE) as never, vi.fn());

        invoked.mockResolvedValue({ ok: false, error: "not-found" });
        await runJumpToSource();
        expect(lastToast()).toBe("That card is no longer in the document.");

        invoked.mockResolvedValue({ ok: false, error: "bad-request" });
        await runJumpToSource();
        expect(lastToast()).toBe("CardMirror could not read this cell's source.");
    });

    it("reports each transport failure in its own words", async () => {
        setActiveHot(makeGrid("Perm solves", SOURCE) as never, vi.fn());
        const cases: [string, string][] = [
            ["not-registered", "CardMirror has never run on this machine."],
            ["not-running", "CardMirror is not running."],
            ["timeout", "CardMirror did not answer."],
            ["bad-response", "CardMirror sent something ebb could not read."],
        ];
        for (const [error, message] of cases) {
            invoked.mockRejectedValue(error);
            await runJumpToSource();
            expect(lastToast()).toBe(message);
        }
    });

    it("falls back to a readable message on an unnamed failure", async () => {
        setActiveHot(makeGrid("Perm solves", SOURCE) as never, vi.fn());
        invoked.mockRejectedValue(new Error("boom"));
        await runJumpToSource();
        expect(lastToast()).toBe("CardMirror sent something ebb could not read.");
    });
});

describe("send to doc", () => {
    it("sends the selected text with the configured type", async () => {
        setActiveHot(makeGrid("Perm solves") as never, vi.fn());
        useFlowStore.setState({ cardmirrorTextType: "tag" });
        invoked.mockResolvedValue({ ok: true, inserted: true, docTitle: "1AR" });

        await runSendToDoc();

        expect(invoked).toHaveBeenCalledWith("cardmirror_insert", {
            text: "Perm solves",
            role: "tag",
            newParagraph: true,
        });
        expect(lastToast()).toBe('Sent to "1AR".');
    });

    it("sends one line per selected cell and skips blanks", async () => {
        const texts: Record<string, string | null> = {
            "0,0": "First",
            "1,0": "   ",
            "2,0": "Second",
        };
        setActiveHot(
            {
                getSelectedLast: () => [0, 0],
                getSelectedRange: () => [
                    {
                        getTopLeftCorner: () => ({ row: 0, col: 0 }),
                        getBottomRightCorner: () => ({ row: 2, col: 0 }),
                    },
                ],
                getCellMeta: () => ({}),
                getDataAtCell: (r: number, c: number) => texts[`${r},${c}`] ?? null,
            } as never,
            vi.fn(),
        );
        invoked.mockResolvedValue({ ok: true, inserted: true });

        await runSendToDoc();

        // One newline per cell: CardMirror makes one block per line, so a
        // blank line here would land an empty paragraph in the document.
        expect(invoked.mock.calls[0][1]).toMatchObject({ text: "First\nSecond" });
        expect(lastToast()).toBe("Sent to CardMirror.");
    });

    it("asks for text before sending an empty selection", async () => {
        setActiveHot(makeGrid(null) as never, vi.fn());
        await runSendToDoc();
        expect(invoked).not.toHaveBeenCalled();
        expect(lastToast()).toBe("Select a cell with text to send.");
    });

    it("explains every insert refusal", async () => {
        setActiveHot(makeGrid("Perm solves") as never, vi.fn());
        const cases: [string, string][] = [
            ["no-target-doc", "Open a document in CardMirror first."],
            ["doc-readonly", "That CardMirror document is in read mode."],
            ["bad-request", "CardMirror would not take that text."],
            ["internal", "CardMirror could not take that text."],
        ];
        for (const [error, message] of cases) {
            invoked.mockResolvedValue({ ok: false, error });
            await runSendToDoc();
            expect(lastToast()).toBe(message);
        }
    });
});

describe("the desktop-only gate", () => {
    it("makes both commands silent no-ops when the switch is off", async () => {
        setActiveHot(makeGrid("Perm solves", SOURCE) as never, vi.fn());
        useFlowStore.setState({ cardmirrorEnabled: false });

        await runJumpToSource();
        await runSendToDoc();

        expect(invoked).not.toHaveBeenCalled();
        expect(toasted).not.toHaveBeenCalled();
    });

    it("makes both commands silent no-ops on the web build", async () => {
        setActiveHot(makeGrid("Perm solves", SOURCE) as never, vi.fn());
        delete (window as unknown as Record<string, unknown>).__TAURI_INTERNALS__;

        await runJumpToSource();
        await runSendToDoc();

        expect(invoked).not.toHaveBeenCalled();
        expect(toasted).not.toHaveBeenCalled();
    });
});

describe("CardMirror's consent gate", () => {
    const WAITING = "Waiting for approval in CardMirror. No need to try again.";

    beforeEach(() => {
        setActiveHot(makeGrid("Perm solves", SOURCE) as never, vi.fn());
    });

    it("waits instead of claiming a send that has not happened yet", async () => {
        // A queued insert answers ok:true with inserted:false - accepted for
        // delivery, nothing written. Reporting that as "Sent" would be a lie.
        invoked.mockResolvedValue({ ok: true, inserted: false, pending: "consent" });

        await runSendToDoc();

        expect(lastToast()).toBe(WAITING);
    });

    it("says a queued jump is waiting rather than staying silent", async () => {
        invoked.mockResolvedValue({ ok: true, jumped: false, pending: "consent" });

        await runJumpToSource();

        expect(lastToast()).toBe(WAITING);
    });

    it("leaves a queued action queued, since approving replays it", async () => {
        invoked.mockResolvedValue({ ok: true, inserted: false, pending: "consent" });

        await runSendToDoc();

        expect(invoked).toHaveBeenCalledTimes(1);
    });

    it("treats an unknown pending kind as unfinished, not as success", async () => {
        invoked.mockResolvedValue({ ok: true, inserted: false, pending: "something-new" });

        await runSendToDoc();

        expect(lastToast()).not.toBe("Sent to CardMirror.");
    });

    it("explains every consent refusal, on both routes alike", async () => {
        const cases: [string, string][] = [
            ["unidentified", "CardMirror did not recognize ebb. Check for an ebb update."],
            ["inserts-disabled", "CardMirror is refusing inserts from other apps."],
            [
                "not-allowed",
                "CardMirror is blocking ebb. Allow it under External apps in its settings.",
            ],
        ];
        for (const [error, message] of cases) {
            invoked.mockResolvedValue({ ok: false, error });

            await runSendToDoc();
            expect(lastToast(), `insert ${error}`).toBe(message);

            await runJumpToSource();
            expect(lastToast(), `jump ${error}`).toBe(message);
        }
    });

    it("never routes around a refusal with a second attempt", async () => {
        invoked.mockResolvedValue({ ok: false, error: "not-allowed" });

        await runSendToDoc();
        await runJumpToSource();

        // One call per command: a denial is the user's decision, so there is
        // no retry and no other way to the document.
        expect(invoked).toHaveBeenCalledTimes(2);
    });
});
