/**
 * PrintView component tests.
 *
 * Verifies that PrintView renders every sheet's title and every data row
 * (never a virtualized subset), with decorations mapped from cell meta, and
 * that the printed RFD carries every author.
 */

import { render, screen } from "@testing-library/react";
import { describe, it, expect, beforeEach } from "vitest";

import PrintView from "@/components/flow/PrintView";
import type { Contacts } from "@/lib/collab/contacts";
import { makeFlowRound } from "@/lib/model/flow";
import { useFlowStore } from "@/lib/store/useFlowStore";

function resetStore() {
    useFlowStore.setState({ round: null, activeSheetId: null, contacts: {} });
}

function setup() {
    const round = makeFlowRound({});
    const flow = round.sheets.find((s) => s.kind !== "cx")!;
    flow.title = "Case";
    flow.data = Array.from({ length: 60 }, (_, r) => [`arg ${r}`, null]);
    flow.meta = { "0,0": { bold: true }, "1,0": { highlight: true } };
    useFlowStore.getState().loadRound(round);
    return round;
}

describe("PrintView", () => {
    beforeEach(resetStore);

    it("renders nothing without a round", () => {
        const { container } = render(<PrintView />);
        expect(container.firstChild).toBeNull();
    });

    it("renders every sheet in order with all data rows", () => {
        const round = setup();
        render(<PrintView />);
        const flow = round.sheets.find((s) => s.kind !== "cx")!;
        const cx = round.sheets.find((s) => s.kind === "cx")!;
        expect(screen.getByTestId(`print-sheet-title-${cx.id}`)).toHaveTextContent("CX");
        expect(screen.getByTestId(`print-sheet-title-${flow.id}`)).toHaveTextContent("Case");
        // All 60 rows render - print never virtualizes.
        expect(screen.getByText("arg 0")).toBeInTheDocument();
        expect(screen.getByText("arg 59")).toBeInTheDocument();
    });

    it("maps cell meta onto decoration classes", () => {
        setup();
        render(<PrintView />);
        expect(screen.getByText("arg 0")).toHaveClass("flow-bold");
        expect(screen.getByText("arg 1")).toHaveClass("flow-highlight");
    });

    it("renders CX period labels in the CX header", () => {
        const round = setup();
        const cx = round.sheets.find((s) => s.kind === "cx")!;
        cx.data = [["q", "a", null, null, null, null, null, null]];
        useFlowStore.getState().loadRound(round);
        render(<PrintView />);
        expect(screen.getAllByText("1AC CX Question")).toHaveLength(1);
    });

    describe("RFD", () => {
        const RAE = "aaa11111aaa";
        const SAM = "bbb22222bbb";
        const contacts: Contacts = {
            [RAE]: { name: "Rae", role: "coach" },
            [SAM]: { name: "Sam", role: "partner" },
        };

        function withDecision(decision: object, table: Contacts = contacts) {
            const round = setup();
            round.scouting.decision = decision;
            useFlowStore.getState().loadRound(round);
            useFlowStore.setState({ contacts: table });
        }

        it("omits the RFD block when there are no notes at all", () => {
            withDecision({ vote: "aff" });
            render(<PrintView />);
            expect(screen.queryByTestId("print-rfd")).toBeNull();
        });

        it("prints the owner's notes and every peer, in the preview's order", () => {
            withDecision({
                rfd: "## Voter\n\nmy own voter",
                peerNotes: { [SAM]: "neg on case", [RAE]: "aff on T" },
            });
            render(<PrintView />);

            const block = screen.getByTestId("print-rfd");
            expect(block.textContent).toContain("my own voter");
            expect(block.querySelector("h2")).not.toBeNull();

            const sections = screen.getAllByTestId("print-rfd-peer-note");
            expect(sections.map((s) => s.querySelector("h3")!.textContent)).toEqual(["Rae", "Sam"]);
            expect(sections[0].textContent).toContain("aff on T");
            expect(sections[1].textContent).toContain("neg on case");
        });

        it("prints peer notes even when the owner wrote none", () => {
            withDecision({ peerNotes: { [RAE]: "aff on T" } });
            render(<PrintView />);
            const sections = screen.getAllByTestId("print-rfd-peer-note");
            expect(sections).toHaveLength(1);
            expect(sections[0].textContent).toContain("aff on T");
        });

        it("names an unknown peer by the short form of its EndpointId", () => {
            withDecision({ peerNotes: { "0123456789abcdef": "dropped the disad" } }, {});
            render(<PrintView />);
            expect(screen.getByTestId("print-rfd-peer-note").textContent).toContain("01234567");
        });

        it("strips a script tag out of a printed peer note", () => {
            withDecision({ peerNotes: { [RAE]: "harmless <script>window.pwned = 1</script>" } });
            render(<PrintView />);
            const section = screen.getByTestId("print-rfd-peer-note");
            expect(section.querySelector("script")).toBeNull();
            expect(section.textContent).not.toContain("pwned");
        });

        it("keeps an unterminated construct inside its own author's section", () => {
            withDecision({
                rfd: "my own voter",
                peerNotes: { [RAE]: "```\nfenced and never closed", [SAM]: "sams own voter" },
            });
            render(<PrintView />);

            const sections = screen.getAllByTestId("print-rfd-peer-note");
            expect(sections[0].textContent).not.toContain("sams own voter");
            expect(sections[1].textContent).toContain("sams own voter");
            expect(screen.getByTestId("print-rfd").textContent).toContain("my own voter");
        });

        it("prints the RFD after the sheets", () => {
            withDecision({ rfd: "my own voter" });
            render(<PrintView />);
            const view = screen.getByTestId("print-view");
            const block = screen.getByTestId("print-rfd");
            expect(view.lastElementChild).toBe(block);
        });
    });
});
