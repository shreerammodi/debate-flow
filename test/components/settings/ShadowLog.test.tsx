import { render, screen, within } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { beforeEach, describe, expect, it } from "vitest";

import ShadowLog from "@/components/settings/ShadowLog";
import type { ShadowEntry } from "@/lib/collab/shadow";
import { columnsForFlowSheet } from "@/lib/grid/flowColumns";
import { makeFlowRound } from "@/lib/model/flow";
import { useCollabStore } from "@/lib/store/useCollabStore";
import { useFlowStore } from "@/lib/store/useFlowStore";

const ALEX = "k51qzi5uqu5dlalex";
const AT = Date.parse("2026-07-26T14:03:00Z");

const round = makeFlowRound({});
const sheet = round.sheets.find((s) => s.kind !== "cx")!;
const firstColumn = columnsForFlowSheet(round, sheet)[0];

function entry(patch: Partial<ShadowEntry> = {}): ShadowEntry {
    return {
        at: AT,
        from: ALEX,
        diffs: [{ sheetId: sheet.id, col: 0, row: 2, mine: "perm", theirs: "perm do both" }],
        dropped: [],
        ...patch,
    };
}

beforeEach(() => {
    useCollabStore.getState().clearShadow();
    useFlowStore.setState({
        round,
        shadowMode: true,
        contacts: { [ALEX]: { name: "Alex", role: "partner" } },
    });
});

describe("ShadowLog", () => {
    it("stays out of the way when shadow mode is off and nothing was recorded", () => {
        useFlowStore.setState({ shadowMode: false });
        render(<ShadowLog />);
        expect(screen.queryByTestId("shadow-log")).toBeNull();
    });

    it("says so plainly when shadow mode is on but nothing has arrived", () => {
        render(<ShadowLog />);
        expect(screen.getByTestId("shadow-log-empty")).toBeTruthy();
        expect(screen.queryByTestId("shadow-log-entry")).toBeNull();
    });

    it("stays readable after shadow mode goes back off, so a record is never stranded", () => {
        useCollabStore.getState().pushShadow(entry());
        useFlowStore.setState({ shadowMode: false });
        render(<ShadowLog />);
        expect(screen.getAllByTestId("shadow-log-entry")).toHaveLength(1);
    });

    it("shows the time, the peer, and the cell that would have changed", () => {
        useCollabStore.getState().pushShadow(entry());
        render(<ShadowLog />);

        const row = screen.getByTestId("shadow-log-entry");
        expect(within(row).getByText(new Date(AT).toLocaleTimeString())).toBeTruthy();
        expect(within(row).getByText(/Alex/)).toBeTruthy();
        const diff = within(row).getByTestId("shadow-log-diff");
        expect(diff.textContent).toContain(sheet.title);
        expect(diff.textContent).toContain(firstColumn.name);
        expect(diff.textContent).toContain("row 3");
        expect(diff.textContent).toContain("perm");
        expect(diff.textContent).toContain("perm do both");
    });

    it("calls a blank cell blank rather than showing nothing", () => {
        useCollabStore
            .getState()
            .pushShadow(
                entry({ diffs: [{ sheetId: sheet.id, col: 0, row: 0, mine: "", theirs: "T" }] }),
            );
        render(<ShadowLog />);
        expect(screen.getByTestId("shadow-log-diff").textContent).toContain("(blank)");
    });

    it("names an unknown peer and an unknown sheet by their ids", () => {
        useFlowStore.setState({ round: null, contacts: {} });
        useCollabStore.getState().pushShadow(entry());
        render(<ShadowLog />);

        const row = screen.getByTestId("shadow-log-entry");
        expect(within(row).getByText(new RegExp(ALEX.slice(0, 8)))).toBeTruthy();
        expect(within(row).getByTestId("shadow-log-diff").textContent).toContain(sheet.id);
    });

    it("lists the newest observation first", () => {
        useCollabStore.getState().pushShadow(entry({ from: "older", at: AT }));
        useCollabStore.getState().pushShadow(entry({ from: "newer", at: AT + 1000 }));
        render(<ShadowLog />);

        const rows = screen.getAllByTestId("shadow-log-entry");
        expect(rows[0].textContent).toContain("newer");
        expect(rows[1].textContent).toContain("older");
    });

    it("reports the cells a merge would have buried", () => {
        useCollabStore.getState().pushShadow(
            entry({
                dropped: [
                    {
                        sheetId: sheet.id,
                        col: 0,
                        rank: "a0",
                        text: "extend the perm",
                        writtenBy: "me",
                        deletedBy: ALEX,
                    },
                ],
            }),
        );
        render(<ShadowLog />);

        const dropped = screen.getByTestId("shadow-log-dropped");
        expect(dropped.textContent).toContain("extend the perm");
        expect(dropped.textContent).toContain("Alex");
    });

    it("keeps a quiet observation and says the link changed nothing", () => {
        useCollabStore.getState().pushShadow(entry({ diffs: [], dropped: [] }));
        render(<ShadowLog />);

        const row = screen.getByTestId("shadow-log-entry");
        expect(row.dataset.state).toBe("quiet");
        expect(row.textContent).toContain("No change");
    });

    it("weighs a buried cell above a plain change", () => {
        useCollabStore.getState().pushShadow(entry());
        useCollabStore.getState().pushShadow(
            entry({
                dropped: [
                    {
                        sheetId: sheet.id,
                        col: 0,
                        rank: "a0",
                        text: "extend the perm",
                        writtenBy: "me",
                        deletedBy: ALEX,
                    },
                ],
            }),
        );
        render(<ShadowLog />);

        const rows = screen.getAllByTestId("shadow-log-entry");
        expect(rows[0].dataset.state).toBe("buried");
        expect(rows[1].dataset.state).toBe("changed");
    });

    it("clears the log on request", async () => {
        const user = userEvent.setup();
        useCollabStore.getState().pushShadow(entry());
        render(<ShadowLog />);

        await user.click(screen.getByTestId("shadow-log-clear"));

        expect(useCollabStore.getState().shadowLog).toEqual([]);
        expect(screen.getByTestId("shadow-log-empty")).toBeTruthy();
    });
});
