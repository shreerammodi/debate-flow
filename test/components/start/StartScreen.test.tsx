import { render, screen, waitFor } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { beforeEach, describe, expect, it, vi } from "vitest";

import { makeFlowRound } from "@/lib/model/flow";
import { serializeFlow } from "@/lib/persistence/flowFile";
import { serializeRecents } from "@/lib/persistence/recents";
import { useFlowStore } from "@/lib/store/useFlowStore";

import { FLOWS_DIR, HOME, installFakeFlowFs, type FakeFlowFs } from "../../support/fakeFlowFs";

const push = vi.fn();
const stableRouter = { push };
vi.mock("next/navigation", () => ({ useRouter: () => stableRouter }));

import StartScreen from "@/components/start/StartScreen";

let fs: FakeFlowFs;

/** Put a readable flow on the fake disk and list it as recent. */
function seedRecents(paths: string[]) {
    for (const path of paths) {
        const round = makeFlowRound({});
        round.scouting.affSchool = "Westwood";
        round.scouting.aff.first = { first: "Ada", last: "Gray" };
        round.scouting.negSchool = "Harvard";
        round.scouting.neg.first = { first: "Ben", last: "Stone" };
        round.scouting.tournament = "Berkeley";
        fs.files.set(path, serializeFlow(round));
    }
    fs.files.set(
        "/config/recents.json",
        serializeRecents(paths.map((path, i) => ({ path, openedAt: 100 - i }))),
    );
}

beforeEach(() => {
    fs = installFakeFlowFs();
    push.mockReset();
    localStorage.setItem("ebb-idb-migrated", "1");
    useFlowStore.setState({ newFlowOpen: false, settingsOpen: false });
});

describe("StartScreen", () => {
    it("never sweeps the old storage without being asked", async () => {
        // The marker is cleared so the legacy check actually looks; with no old
        // database present nothing should be written and nothing prompted.
        localStorage.removeItem("ebb-idb-migrated");
        render(<StartScreen />);
        await waitFor(() => expect(screen.getByTestId("start-new")).toBeInTheDocument());
        expect(fs.writes).toEqual([]);
        expect(screen.queryByTestId("migration-dialog")).not.toBeInTheDocument();
    });

    it("offers the three commands", async () => {
        render(<StartScreen />);
        expect(screen.getByTestId("start-new")).toHaveTextContent("New flow");
        expect(screen.getByTestId("start-open")).toHaveTextContent("Open");
        expect(screen.getByTestId("start-settings")).toHaveTextContent("Settings");
        await waitFor(() => expect(screen.queryByTestId("start-recent-1")).not.toBeInTheDocument());
    });

    it("labels a recent flow by its matchup and shortens the path", async () => {
        seedRecents([`${FLOWS_DIR}/berkeley.ebb`]);
        render(<StartScreen />);

        const row = await screen.findByTestId("start-recent-1");
        expect(row).toHaveTextContent("Westwood AG vs Harvard BS");
        expect(row).toHaveTextContent("Berkeley");
        expect(row).toHaveTextContent(`${FLOWS_DIR.replace(HOME, "~")}/berkeley.ebb`);
    });

    it("falls back to the filename when a flow will not parse", async () => {
        fs.files.set("/a/broken.ebb", "{ truncated");
        fs.files.set(
            "/config/recents.json",
            serializeRecents([{ path: "/a/broken.ebb", openedAt: 1 }]),
        );

        render(<StartScreen />);

        // The row survives so the user can open it and be told what is wrong.
        expect(await screen.findByTestId("start-recent-1")).toHaveTextContent("broken");
    });

    it("drops a recent whose file is gone", async () => {
        fs.files.set(
            "/config/recents.json",
            serializeRecents([{ path: "/a/vanished.ebb", openedAt: 1 }]),
        );

        render(<StartScreen />);

        await waitFor(() => {
            expect(fs.files.get("/config/recents.json")).not.toContain("vanished");
        });
        expect(screen.queryByTestId("start-recent-1")).not.toBeInTheDocument();
    });

    it("shows at most six recents, each on its own number key", async () => {
        seedRecents(Array.from({ length: 9 }, (_, i) => `${FLOWS_DIR}/r${i}.ebb`));
        render(<StartScreen />);

        await screen.findByTestId("start-recent-1");
        expect(screen.getByTestId("start-recent-6")).toBeInTheDocument();
        expect(screen.queryByTestId("start-recent-7")).not.toBeInTheDocument();
    });

    it("opens a recent by its number key", async () => {
        seedRecents([`${FLOWS_DIR}/a.ebb`, `${FLOWS_DIR}/b.ebb`]);
        render(<StartScreen />);
        await screen.findByTestId("start-recent-2");

        await userEvent.keyboard("2");

        expect(push).toHaveBeenCalledWith(`/flow?path=${encodeURIComponent(`${FLOWS_DIR}/b.ebb`)}`);
    });

    it("opens the New flow prompt on n", async () => {
        render(<StartScreen />);
        await userEvent.keyboard("n");
        expect(useFlowStore.getState().newFlowOpen).toBe(true);
    });

    it("opens Settings on s", async () => {
        render(<StartScreen />);
        await userEvent.keyboard("s");
        expect(useFlowStore.getState().settingsOpen).toBe(true);
    });

    it("walks the column with j and opens the highlighted row with Enter", async () => {
        seedRecents([`${FLOWS_DIR}/a.ebb`]);
        render(<StartScreen />);
        await screen.findByTestId("start-recent-1");

        // Three actions, then the single recent.
        await userEvent.keyboard("jjj{Enter}");

        expect(push).toHaveBeenCalledWith(`/flow?path=${encodeURIComponent(`${FLOWS_DIR}/a.ebb`)}`);
    });

    it("forgets a recent without deleting its file", async () => {
        const path = `${FLOWS_DIR}/a.ebb`;
        seedRecents([path]);
        render(<StartScreen />);
        await screen.findByTestId("start-recent-1");

        await userEvent.click(screen.getByRole("button", { name: /Remove .* from recent flows/ }));

        expect(screen.queryByTestId("start-recent-1")).not.toBeInTheDocument();
        expect(fs.files.has(path)).toBe(true);
    });

    it("links out to the docs, the repo, and the author", () => {
        render(<StartScreen />);
        expect(screen.getByRole("link", { name: "Documentation" })).toHaveAttribute(
            "href",
            "https://ebb.smodi.net/docs",
        );
        expect(screen.getByRole("link", { name: "GitHub" })).toHaveAttribute(
            "href",
            "https://github.com/shreerammodi/ebb",
        );
        // Only the name is the link; "Developed by" is plain text beside it.
        expect(screen.getByRole("link", { name: "Shreeram Modi" })).toHaveAttribute(
            "href",
            "https://smodi.net",
        );
    });
});
