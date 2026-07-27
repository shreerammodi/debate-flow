import { EditorView } from "@codemirror/view";
import { render, screen } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { describe, it, expect, beforeEach, vi } from "vitest";

import RfdDrawer from "@/components/flow/RfdDrawer";
import { TooltipProvider } from "@/components/ui/tooltip";
import type { Contacts } from "@/lib/collab/contacts";
import { focusActiveHot } from "@/lib/grid/hotInstance";
import { makeFlowRound } from "@/lib/model/flow";
import { useFlowStore } from "@/lib/store/useFlowStore";

vi.mock("@/lib/grid/hotInstance", () => ({ focusActiveHot: vi.fn() }));

function renderDrawer() {
    return render(
        <TooltipProvider>
            <RfdDrawer />
        </TooltipProvider>,
    );
}

describe("RfdDrawer", () => {
    beforeEach(() => {
        const round = makeFlowRound({});
        round.scouting.decision = { rfd: "aff on T" };
        useFlowStore.getState().loadRound(round);
        useFlowStore.getState().setRfdOpen(true);
        useFlowStore.getState().setRfdVim(false);
        useFlowStore.setState({ contacts: {} });
    });

    it("mounts a CodeMirror editor seeded with the stored RFD", () => {
        const { container } = renderDrawer();
        expect(screen.getByTestId("rfd-drawer")).toBeInTheDocument();
        expect(container.querySelector(".cm-editor")).not.toBeNull();
        expect(container.textContent).toContain("aff on T");
    });

    it("closes the drawer when the close button is clicked", async () => {
        renderDrawer();
        await userEvent.click(screen.getByTestId("rfd-close"));
        expect(useFlowStore.getState().rfdOpen).toBe(false);
    });

    it("returns focus to the grid when it unmounts", () => {
        vi.mocked(focusActiveHot).mockClear();
        const { unmount } = renderDrawer();
        unmount();
        expect(focusActiveHot).toHaveBeenCalled();
    });

    it("shows the vim status bar when rfdVim is enabled", () => {
        useFlowStore.getState().setRfdVim(true);
        const { container } = renderDrawer();
        expect(container.querySelector(".cm-vim-panel")).not.toBeNull();
    });

    it("renders the RFD as markdown in preview mode", async () => {
        const round = makeFlowRound({});
        round.scouting.decision = { rfd: "## Voter\n\n- aff on T" };
        useFlowStore.getState().loadRound(round);
        useFlowStore.getState().setRfdOpen(true);

        renderDrawer();
        await userEvent.click(screen.getByTestId("rfd-preview-toggle"));

        const preview = screen.getByTestId("rfd-preview");
        expect(preview.querySelector("h2")).not.toBeNull();
        expect(preview.querySelector("li")).not.toBeNull();
        expect(preview.textContent).toContain("Voter");
    });

    describe("peer notes", () => {
        const RAE = "aaa11111aaa";
        const SAM = "bbb22222bbb";

        function loadPeers(peerNotes: Record<string, string>, contacts: Contacts = {}) {
            const round = makeFlowRound({});
            round.scouting.decision = { rfd: "my own voter", peerNotes };
            useFlowStore.getState().loadRound(round);
            useFlowStore.getState().setRfdOpen(true);
            useFlowStore.setState({ contacts });
        }

        async function openPreview() {
            const view = renderDrawer();
            await userEvent.click(screen.getByTestId("rfd-preview-toggle"));
            return view;
        }

        it("renders one section per peer below the owner's own notes", async () => {
            loadPeers(
                { [SAM]: "neg on case", [RAE]: "aff on T" },
                {
                    [RAE]: { name: "Rae", role: "coach" },
                    [SAM]: { name: "Sam", role: "partner" },
                },
            );
            await openPreview();

            const sections = screen.getAllByTestId("rfd-peer-note");
            expect(sections).toHaveLength(2);
            expect(sections[0].textContent).toContain("Rae");
            expect(sections[0].textContent).toContain("aff on T");
            expect(sections[1].textContent).toContain("Sam");
            expect(sections[1].textContent).toContain("neg on case");

            const preview = screen.getByTestId("rfd-preview");
            const peers = screen.getByTestId("rfd-peer-notes");
            expect(preview.textContent).toContain("my own voter");
            expect(peers.textContent).not.toContain("my own voter");
            expect(preview.compareDocumentPosition(peers)).toBe(
                Node.DOCUMENT_POSITION_CONTAINED_BY | Node.DOCUMENT_POSITION_FOLLOWING,
            );
        });

        it("names an unknown peer by the short form of its EndpointId", async () => {
            loadPeers({ "0123456789abcdef": "dropped the disad" });
            await openPreview();
            expect(screen.getByTestId("rfd-peer-note").textContent).toContain("01234567");
        });

        it("renders a peer's markdown", async () => {
            loadPeers({ [RAE]: "## Voter\n\n- aff on T" });
            await openPreview();
            const section = screen.getByTestId("rfd-peer-note");
            expect(section.querySelector("h2")).not.toBeNull();
            expect(section.querySelector("li")).not.toBeNull();
        });

        it("strips a script tag out of a peer's note", async () => {
            loadPeers({ [RAE]: "harmless <script>window.pwned = 1</script> tail" });
            await openPreview();
            const section = screen.getByTestId("rfd-peer-note");
            expect(section.querySelector("script")).toBeNull();
            expect(section.innerHTML).not.toContain("<script");
            expect(section.textContent).not.toContain("pwned");
            expect(section.textContent).toContain("harmless");
        });

        it("strips an inline event handler out of a peer's note", async () => {
            loadPeers({ [RAE]: '<img src="x" onerror="window.pwned = 1">' });
            await openPreview();
            const img = screen.getByTestId("rfd-peer-note").querySelector("img");
            expect(img?.getAttribute("onerror")).toBeNull();
        });

        it("keeps an unterminated construct inside its own author's section", async () => {
            loadPeers({ [RAE]: "```\nfenced and never closed", [SAM]: "sams own voter" });
            await openPreview();

            const sections = screen.getAllByTestId("rfd-peer-note");
            expect(sections[0].textContent).toContain("fenced and never closed");
            expect(sections[0].textContent).not.toContain("sams own voter");
            expect(sections[1].textContent).toContain("sams own voter");
            expect(sections[1].querySelector("code")).toBeNull();
            expect(screen.getByTestId("rfd-preview").textContent).toContain("my own voter");
        });

        it("keeps an unclosed tag from swallowing the next author", async () => {
            loadPeers({ [RAE]: "<div>never closed", [SAM]: "sams own voter" });
            await openPreview();

            const sections = screen.getAllByTestId("rfd-peer-note");
            expect(sections[0].textContent).not.toContain("sams own voter");
            expect(sections[1].textContent).toContain("sams own voter");
        });

        it("omits the peer block entirely when no peer has notes", async () => {
            loadPeers({ [RAE]: "   " });
            await openPreview();
            expect(screen.queryByTestId("rfd-peer-notes")).toBeNull();
            expect(screen.getByTestId("rfd-preview").textContent).toContain("my own voter");
        });

        it("keeps peer text out of the edit pane", async () => {
            loadPeers({ [RAE]: "a peer wrote this" }, { [RAE]: { name: "Rae", role: "coach" } });
            const { container } = renderDrawer();

            const editor = container.querySelector(".cm-content")!;
            expect(editor.textContent).toBe("my own voter");
            expect(container.textContent).not.toContain("a peer wrote this");
            expect(screen.queryByTestId("rfd-peer-notes")).toBeNull();

            await userEvent.click(screen.getByTestId("rfd-preview-toggle"));
            expect(screen.getByTestId("rfd-peer-note").textContent).toContain("a peer wrote this");
        });

        it("leaves peerNotes untouched when the owner edits their own notes", () => {
            loadPeers({ [RAE]: "a peer wrote this" });
            const { container } = renderDrawer();

            const view = EditorView.findFromDOM(container.querySelector(".cm-editor")!)!;
            view.dispatch({ changes: { from: 0, to: view.state.doc.length, insert: "rewritten" } });

            const decision = useFlowStore.getState().round!.scouting.decision!;
            expect(decision.rfd).toBe("rewritten");
            expect(decision.peerNotes).toEqual({ [RAE]: "a peer wrote this" });
        });
    });
});
