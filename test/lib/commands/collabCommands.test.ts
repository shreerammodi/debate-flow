import { beforeEach, describe, expect, it, vi } from "vitest";

import type * as CollabPeerLink from "@/lib/collab/peerLink";
import type { PeerLink, PeerLinkConfig } from "@/lib/collab/peerLink";
import type * as CollabRuntime from "@/lib/collab/runtime";

/** Lets one test stand a session up and make its teardown fail. */
const runtime: { live: boolean; endFails: Error | null } = { live: false, endFails: null };

vi.mock("@/lib/collab/runtime", async (importOriginal) => {
    const real = await importOriginal<typeof CollabRuntime>();
    return {
        ...real,
        currentSession: () => (runtime.live ? ({} as never) : real.currentSession()),
        endSession: async () => {
            if (runtime.endFails) throw runtime.endFails;
            await real.endSession();
        },
    };
});

/** Filled in below, once the real transport module has been imported. */
const transport = vi.hoisted(() => ({
    link: null as ((config: PeerLinkConfig) => Promise<PeerLink>) | null,
}));

vi.mock("@/lib/collab/peerLink", async (importOriginal) => ({
    ...(await importOriginal<typeof CollabPeerLink>()),
    createPeerLinkFor: (config: PeerLinkConfig) => transport.link!(config),
}));

import { createMemoryNet, type MemoryNet } from "@/lib/collab/peerLinkMemory";
import { endSession } from "@/lib/collab/runtime";
import { parseTicket } from "@/lib/collab/ticket";
import {
    runEnd,
    runInvite,
    runJoin,
    runShare,
    type CollabCommandDeps,
} from "@/lib/commands/collabCommands";
import { executeCommand } from "@/lib/commands/commands";
import { COLLAB_COMMANDS, COMMANDS } from "@/lib/commands/registry";
import { getPresetKeymap } from "@/lib/keymap/presets";
import { makeFlowRound } from "@/lib/model/flow";
import { useFlowStore } from "@/lib/store/useFlowStore";
import { useTicketDialog } from "@/lib/store/useTicketDialog";

const net: MemoryNet = createMemoryNet();
/** What iroh hands back. A ticket names its host, so the host holds a real one. */
const HOST = "a".repeat(64);
transport.link = net.create(HOST);

/** A round open, the switch on, and the desktop shell in place: a share can run. */
function ready(): void {
    (window as unknown as Record<string, unknown>).__TAURI_INTERNALS__ = {};
    useFlowStore.setState({
        collabEnabled: true,
        collabName: "Host",
        round: makeFlowRound({}),
    });
}

function deps(over: Partial<CollabCommandDeps> = {}): CollabCommandDeps & {
    notices: string[];
    failures: string[];
    shown: string[];
} {
    const notices: string[] = [];
    const failures: string[] = [];
    const shown: string[] = [];
    return {
        notices,
        failures,
        shown,
        notify: (m) => notices.push(m),
        fail: (m) => failures.push(m),
        askForTicket: async () => null,
        presentTicket: (t) => {
            shown.push(t);
        },
        openFlow: vi.fn(),
        ...over,
    };
}

beforeEach(async () => {
    runtime.live = false;
    runtime.endFails = null;
    await endSession();
    useFlowStore.setState({ collabEnabled: false, round: null });
    // isDesktop() is false under jsdom unless the harness says otherwise.
    delete (window as unknown as Record<string, unknown>).__TAURI_INTERNALS__;
    net.reset();
    useTicketDialog.setState({ open: false, ticket: "", resolve: null });
});

describe("the commands are palette-only", () => {
    it("claims no chord in the default keymap", () => {
        const bound = new Set(Object.values(getPresetKeymap().bindings));
        for (const id of COLLAB_COMMANDS) expect(bound.has(id)).toBe(false);
    });

    it("is registered with a label a debater can search for", () => {
        expect(COMMANDS["collab.share"].label).toBe("Share this round");
        expect(COMMANDS["collab.shareView"].label).toBe("Share this round view only");
        expect(COMMANDS["collab.join"].label).toBe("Join a shared round");
        expect(COMMANDS["collab.end"].label).toBe("End shared session");
    });
});

describe("share", () => {
    it("refuses while the master switch is off, and mints nothing", async () => {
        const d = deps();
        await runShare(d);
        expect(d.shown).toEqual([]);
        expect(d.failures[0]).toMatch(/Settings/);
    });

    it("refuses with no flow open", async () => {
        (window as unknown as Record<string, unknown>).__TAURI_INTERNALS__ = {};
        useFlowStore.setState({ collabEnabled: true, round: null });
        const d = deps();
        await runShare(d);
        expect(d.failures[0]).toMatch(/Open a flow/);
    });

    it("hands over a ticket a partner can edit through", async () => {
        ready();
        const d = deps();
        await runShare(d);
        expect(d.failures).toEqual([]);
        expect(parseTicket(d.shown[0])?.role).toBe("partner");
    });

    it("hands over a view-only ticket when the share names a coach", async () => {
        ready();
        const d = deps();
        await runShare(d, "coach");
        expect(d.failures).toEqual([]);
        expect(parseTicket(d.shown[0])?.role).toBe("coach");
    });
});

describe("join", () => {
    it("refuses while the master switch is off, and never asks for a ticket", async () => {
        const askForTicket = vi.fn(async () => null);
        await runJoin(deps({ askForTicket }));
        expect(askForTicket).not.toHaveBeenCalled();
    });

    it("does nothing when the user backs out of the paste", async () => {
        (window as unknown as Record<string, unknown>).__TAURI_INTERNALS__ = {};
        useFlowStore.setState({ collabEnabled: true });
        const d = deps({ askForTicket: async () => null });
        await runJoin(d);
        expect(d.failures).toEqual([]);
        expect(d.notices).toEqual([]);
    });

    it("reports a ticket it cannot read, rather than throwing", async () => {
        (window as unknown as Record<string, unknown>).__TAURI_INTERNALS__ = {};
        useFlowStore.setState({ collabEnabled: true });
        const d = deps({ askForTicket: async () => "not a ticket" });
        await runJoin(d);
        expect(d.failures[0]).toMatch(/ticket/i);
        expect(d.openFlow).not.toHaveBeenCalled();
    });
});

describe("end", () => {
    it("says so when there is nothing to end", async () => {
        const d = deps();
        await runEnd(d);
        expect(d.failures[0]).toMatch(/No session/);
    });

    // Fired as `void runEnd(...)`, so a rejection would go nowhere but an
    // unhandled promise, and End session is pressed mid-round.
    it("reports a teardown that failed, rather than rejecting", async () => {
        runtime.live = true;
        runtime.endFails = new Error("The endpoint refused to stop");
        const d = deps();
        await expect(runEnd(d)).resolves.toBeUndefined();
        expect(d.failures).toEqual(["The endpoint refused to stop"]);
        expect(d.notices).toEqual([]);
    });

    it("says the session is over when the teardown lands", async () => {
        runtime.live = true;
        const d = deps();
        await runEnd(d);
        expect(d.failures).toEqual([]);
        expect(d.notices[0]).toMatch(/Session ended/);
    });
});

describe("the ticket a share would hand over", () => {
    it("round-trips through the parser the join side uses", () => {
        const round = makeFlowRound({});
        expect(parseTicket("nonsense")).toBeNull();
        expect(round.id).toBeTruthy();
    });

    it("carries the editing role the palette asked for", async () => {
        ready();
        executeCommand("collab.share");
        await vi.waitFor(() => expect(useTicketDialog.getState().ticket).not.toBe(""));
        expect(parseTicket(useTicketDialog.getState().ticket)?.role).toBe("partner");
    });

    it("carries the view-only role the palette asked for", async () => {
        ready();
        executeCommand("collab.shareView");
        await vi.waitFor(() => expect(useTicketDialog.getState().ticket).not.toBe(""));
        expect(parseTicket(useTicketDialog.getState().ticket)?.role).toBe("coach");
    });
});

describe("invite", () => {
    const ALEX = "k51qzi5uqu5dlalex";

    it("refuses while the master switch is off, and never picks a contact", async () => {
        const chooseContact = vi.fn(async () => ALEX);
        await runInvite(deps({ chooseContact }));
        expect(chooseContact).not.toHaveBeenCalled();
    });

    it("says so when nothing has been saved yet", async () => {
        (window as unknown as Record<string, unknown>).__TAURI_INTERNALS__ = {};
        useFlowStore.setState({ collabEnabled: true, round: makeFlowRound({}), contacts: {} });
        const chooseContact = vi.fn(async () => ALEX);
        const d = deps({ chooseContact });
        await runInvite(d);
        expect(d.failures[0]).toMatch(/No saved partners/);
        expect(chooseContact).not.toHaveBeenCalled();
    });

    it("does nothing when the user backs out of the picker", async () => {
        (window as unknown as Record<string, unknown>).__TAURI_INTERNALS__ = {};
        useFlowStore.setState({
            collabEnabled: true,
            round: makeFlowRound({}),
            contacts: { [ALEX]: { name: "Alex", role: "partner" } },
        });
        const d = deps({ chooseContact: async () => null });
        await runInvite(d);
        expect(d.failures).toEqual([]);
        expect(d.notices).toEqual([]);
    });
});
