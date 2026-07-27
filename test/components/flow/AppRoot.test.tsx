import { render, screen, waitFor } from "@testing-library/react";
import { beforeEach, describe, expect, it, vi } from "vitest";

import { TooltipProvider } from "@/components/ui/tooltip";
import { UpdateProvider } from "@/components/update/UpdateProvider";
import { forgetRoundPeers, rememberRoundPeers } from "@/lib/collab/roundPeers";
import { currentSession } from "@/lib/collab/runtime";
import { makeFlowRound } from "@/lib/model/flow";
import { serializeFlow } from "@/lib/persistence/flowFile";
import { useCollabStore } from "@/lib/store/useCollabStore";
import { useFlowStore } from "@/lib/store/useFlowStore";

import { installFakeFlowFs, type FakeFlowFs } from "../../support/fakeFlowFs";

// --- Navigation mock ----------------------------------------------------------

const replace = vi.fn();
let mockSearch = "";

// Stable router object - recreating it each render would change the useEffect
// dependency and cause the effect to re-run indefinitely in tests.
const stableRouter = { replace };

vi.mock("next/navigation", () => ({
    useRouter: () => stableRouter,
    useSearchParams: () => new URLSearchParams(mockSearch),
}));

// Import AppRoot AFTER the mock is set up.
import AppRoot from "@/components/flow/AppRoot";

function mount() {
    return render(
        <TooltipProvider>
            <UpdateProvider>
                <AppRoot />
            </UpdateProvider>
        </TooltipProvider>,
    );
}

const route = (path: string) => `path=${encodeURIComponent(path)}`;

let fs: FakeFlowFs;

beforeEach(() => {
    fs = installFakeFlowFs();
    mockSearch = "";
    replace.mockReset();
    useFlowStore.setState({ round: null, docPath: null, activeSheetId: null });
});

describe("AppRoot", () => {
    it("redirects to the start screen with no ?path=", async () => {
        mount();
        await waitFor(() => expect(replace).toHaveBeenCalledWith("/"));
    });

    it("redirects when the file no longer exists", async () => {
        mockSearch = route("/gone.ebb");
        mount();
        await waitFor(() => expect(replace).toHaveBeenCalledWith("/"));
    });

    it("redirects when the file cannot be parsed", async () => {
        fs.files.set("/broken.ebb", "{ truncated");
        mockSearch = route("/broken.ebb");
        mount();
        await waitFor(() => expect(replace).toHaveBeenCalledWith("/"));
    });

    it("opens the workspace for a readable flow", async () => {
        const round = makeFlowRound({});
        fs.files.set("/a.ebb", serializeFlow(round));
        mockSearch = route("/a.ebb");

        mount();

        await waitFor(() => expect(screen.getByTestId("workspace")).toBeInTheDocument());
        expect(useFlowStore.getState().round?.id).toBe(round.id);
    });

    it("records the path so autosave writes back to the file it opened", async () => {
        const round = makeFlowRound({});
        fs.files.set("/a.ebb", serializeFlow(round));
        mockSearch = route("/a.ebb");

        mount();

        await waitFor(() => expect(useFlowStore.getState().docPath).toBe("/a.ebb"));
    });

    it("adds the flow to recents when it opens", async () => {
        fs.files.set("/a.ebb", serializeFlow(makeFlowRound({})));
        mockSearch = route("/a.ebb");

        mount();

        await waitFor(() => {
            expect(fs.files.get("/config/recents.json")).toContain("/a.ebb");
        });
    });

    it("does not reload a file the store is already editing", async () => {
        const round = makeFlowRound({});
        // Save As leaves the store on the new path and then rewrites the URL;
        // re-reading would only flash the loading frame.
        useFlowStore.setState({ round, docPath: "/a.ebb" });
        mockSearch = route("/a.ebb");

        mount();

        await waitFor(() => expect(screen.getByTestId("workspace")).toBeInTheDocument());
        expect(replace).not.toHaveBeenCalled();
    });
});

/**
 * Shared editing is an iroh endpoint and the browser cannot bind one, so
 * opening a round on the web reaches for no session at all - not one that
 * fails, and not one held by a stand-in transport.
 */
describe("opening a round outside the desktop shell", () => {
    it("does not reach for a session, even for a round with peers", async () => {
        const round = makeFlowRound({});
        fs.files.set("/shared.ebb", serializeFlow(round));
        rememberRoundPeers(round.id, ["a".repeat(64)]);
        useFlowStore.setState({ collabEnabled: true });
        mockSearch = route("/shared.ebb");

        mount();

        await waitFor(() => expect(screen.getByTestId("workspace")).toBeInTheDocument());
        expect(currentSession()).toBeNull();
        // Nothing was attempted, so nothing failed and nothing was reported.
        expect(useCollabStore.getState().status).toBe("off");
        forgetRoundPeers();
    });
});
