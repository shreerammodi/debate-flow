import { beforeEach, describe, expect, it, vi } from "vitest";

import type * as CollabPeerLink from "@/lib/collab/peerLink";
import type { PeerLink, PeerLinkConfig } from "@/lib/collab/peerLink";
import type * as CollabRuntime from "@/lib/collab/runtime";
import type { Role } from "@/lib/collab/types";

/** Lets one test stand a session up, fail its teardown, and read what it dialled. */
const runtime: {
    live: boolean;
    endFails: Error | null;
    invited: { endpointId: string; role: Role }[];
} = { live: false, endFails: null, invited: [] };

vi.mock("@/lib/collab/runtime", async (importOriginal) => {
    const real = await importOriginal<typeof CollabRuntime>();
    return {
        ...real,
        currentSession: () => (runtime.live ? ({} as never) : real.currentSession()),
        endSession: async () => {
            if (runtime.endFails) throw runtime.endFails;
            await real.endSession();
        },
        // Recorded rather than run: what the command owes the runtime is the
        // contact and the grade the picker answered with, and a real dial
        // would decide neither.
        inviteContact: async (_round: unknown, endpointId: string, role: Role) => {
            runtime.invited.push({ endpointId, role });
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
import { useShareSheet } from "@/lib/store/useShareSheet";

const net: MemoryNet = createMemoryNet();
/** What iroh hands back. A ticket names its host, so the host holds a real one. */
const HOST = "a".repeat(64);

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
    codes: string[];
    sheetFailures: string[];
    opened: { role: Role; warning: string }[];
} {
    const notices: string[] = [];
    const failures: string[] = [];
    const codes: string[] = [];
    const sheetFailures: string[] = [];
    const opened: { role: Role; warning: string }[] = [];
    return {
        notices,
        failures,
        codes,
        sheetFailures,
        opened,
        notify: (m) => notices.push(m),
        fail: (m) => failures.push(m),
        // Answered from the switch rather than by a dialog, so a test that
        // wants the refusal turns the switch off and gets it.
        consent: async () => useFlowStore.getState().collabEnabled,
        openShare: (role, warning) => opened.push({ role, warning }),
        showCode: (code) => codes.push(code),
        failShare: (m) => sheetFailures.push(m),
        askForCode: async () => null,
        openFlow: vi.fn(),
        ...over,
    };
}

beforeEach(async () => {
    runtime.live = false;
    runtime.endFails = null;
    runtime.invited.length = 0;
    await endSession();
    useFlowStore.setState({ collabEnabled: false, collabRelayEnabled: true, round: null });
    // isDesktop() is false under jsdom unless the harness says otherwise.
    delete (window as unknown as Record<string, unknown>).__TAURI_INTERNALS__;
    net.reset();
    // One test swaps in a transport that refuses every code, so each starts
    // from the plain one again.
    transport.link = net.create(HOST);
    useShareSheet.setState({ open: false, code: "", screen: "ready" });
});

describe("the commands are palette-only", () => {
    it("claims no chord in the default keymap", () => {
        const bound = new Set(Object.values(getPresetKeymap().bindings));
        for (const id of COLLAB_COMMANDS) expect(bound.has(id)).toBe(false);
    });

    it("is registered with a label a debater can search for", () => {
        expect(COMMANDS["collab.share"].label).toBe("Generate a code to edit");
        expect(COMMANDS["collab.shareView"].label).toBe("Generate a code to view");
        expect(COMMANDS["collab.join"].label).toBe("Join with a code");
        expect(COMMANDS["collab.invite"].label).toBe("Invite a saved partner to edit");
        expect(COMMANDS["collab.inviteView"].label).toBe("Invite a saved partner to view");
        expect(COMMANDS["collab.end"].label).toBe("End shared session");
    });
});

describe("share", () => {
    it("refuses while consent is withheld, and puts no sheet up", async () => {
        const d = deps();
        await runShare(d);
        expect(d.opened).toEqual([]);
        expect(d.codes).toEqual([]);
    });

    it("refuses with no flow open", async () => {
        (window as unknown as Record<string, unknown>).__TAURI_INTERNALS__ = {};
        useFlowStore.setState({ collabEnabled: true, collabRelayEnabled: true, round: null });
        const d = deps();
        await runShare(d);
        expect(d.failures[0]).toMatch(/Open a flow/);
    });

    it("puts the sheet up before it has a code, and fills it in", async () => {
        ready();
        const d = deps();
        await runShare(d);
        expect(d.opened).toEqual([{ role: "editor", warning: "" }]);
        expect(d.codes[0]).toMatch(/^[0-9A-HJKMNP-TV-Z]{8}$/);
        expect(d.sheetFailures).toEqual([]);
    });

    it("names a view-only code for what it grants", async () => {
        ready();
        const d = deps();
        await runShare(d, "viewer");
        expect(d.opened[0].role).toBe("viewer");
        expect(d.codes).toHaveLength(1);
    });

    it("warns before minting that relaying off keeps a round on one wifi", async () => {
        ready();
        useFlowStore.setState({ collabRelayEnabled: false });
        const d = deps();
        await runShare(d);
        expect(d.opened[0].warning).toBe(
            "Relaying is off, so this code only works on the same wifi.",
        );
    });

    it("puts the reason on the sheet, and no code, when pairing fails", async () => {
        ready();
        // Every code's relay refuses, which is what the retry gives up on.
        const base = net.create(HOST);
        transport.link = async (config) => {
            const link = await base(config);
            return {
                ...link,
                async pairHost() {
                    throw new Error("Could not reach the relay for that code");
                },
            };
        };
        const d = deps();
        await runShare(d);
        expect(d.codes).toEqual([]);
        expect(d.sheetFailures[0]).toMatch(/relay/);
    });
});

describe("join", () => {
    it("refuses while consent is withheld, and never asks for a code", async () => {
        const askForCode = vi.fn(async () => null);
        await runJoin(deps({ askForCode }));
        expect(askForCode).not.toHaveBeenCalled();
    });

    it("does nothing when the user backs out of the field", async () => {
        (window as unknown as Record<string, unknown>).__TAURI_INTERNALS__ = {};
        useFlowStore.setState({ collabEnabled: true });
        const d = deps({ askForCode: async () => null });
        await runJoin(d);
        expect(d.failures).toEqual([]);
        expect(d.notices).toEqual([]);
    });

    it("reports a code nobody is holding, rather than throwing", async () => {
        (window as unknown as Record<string, unknown>).__TAURI_INTERNALS__ = {};
        useFlowStore.setState({ collabEnabled: true });
        const d = deps({ askForCode: async () => "TESTAA01" });
        await runJoin(d);
        expect(d.failures).toHaveLength(1);
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

describe("the code a share puts on the air", () => {
    it("reaches the sheet from the palette", async () => {
        ready();
        executeCommand("collab.share");
        await vi.waitFor(() => expect(useShareSheet.getState().code).not.toBe(""));
        expect(useShareSheet.getState().role).toBe("editor");
    });

    it("carries the view-only role the palette asked for", async () => {
        ready();
        executeCommand("collab.shareView");
        await vi.waitFor(() => expect(useShareSheet.getState().code).not.toBe(""));
        expect(useShareSheet.getState().role).toBe("viewer");
    });
});

describe("invite", () => {
    const ALEX = "k51qzi5uqu5dlalex";

    /** A round open, the switch on, and Alex saved: the picker can run. */
    function alexIsSaved(): void {
        (window as unknown as Record<string, unknown>).__TAURI_INTERNALS__ = {};
        useFlowStore.setState({
            collabEnabled: true,
            round: makeFlowRound({}),
            contacts: { [ALEX]: { name: "Alex" } },
        });
    }

    /** Answers the picker the way a click on one of its rows does. */
    const picksAlex = async () => ALEX;

    it("refuses while the master switch is off, and never picks a contact", async () => {
        const chooseContact = vi.fn(picksAlex);
        await runInvite(deps({ chooseContact }));
        expect(chooseContact).not.toHaveBeenCalled();
    });

    it("says so when nothing has been saved yet", async () => {
        (window as unknown as Record<string, unknown>).__TAURI_INTERNALS__ = {};
        useFlowStore.setState({ collabEnabled: true, round: makeFlowRound({}), contacts: {} });
        const chooseContact = vi.fn(picksAlex);
        const d = deps({ chooseContact });
        await runInvite(d);
        expect(d.failures[0]).toMatch(/No saved partners/);
        expect(chooseContact).not.toHaveBeenCalled();
    });

    // Backing out is an answer, and the answer is nobody: a picker that
    // dialled on the way out would put this round on a partner's screen for a
    // gesture that means the opposite.
    it("dials nobody when the user backs out of the picker", async () => {
        alexIsSaved();
        const d = deps({ chooseContact: async () => null });
        await runInvite(d);
        expect(runtime.invited).toEqual([]);
        expect(d.failures).toEqual([]);
        expect(d.notices).toEqual([]);
    });

    // The grade belongs to this invitation, not to the contact row: the same
    // partner is an editor on one round and a viewer on the next. It rides in
    // from the menu entry the debater clicked, and the picker is only asked
    // who, so what the picker is handed is what the runtime gets.
    it("hands the runtime the grade the command was invoked with", async () => {
        alexIsSaved();
        const viewing = deps({ chooseContact: vi.fn(picksAlex) });
        await runInvite(viewing, "viewer");
        expect(runtime.invited).toEqual([{ endpointId: ALEX, role: "viewer" }]);
        expect(viewing.chooseContact).toHaveBeenCalledWith({ [ALEX]: { name: "Alex" } }, "viewer");

        runtime.invited.length = 0;
        await runInvite(deps({ chooseContact: picksAlex }), "editor");
        expect(runtime.invited).toEqual([{ endpointId: ALEX, role: "editor" }]);
    });

    // Edit is never what happens because nobody said otherwise; it is what the
    // menu entry named. The default only spares every caller the argument.
    it("grants edit when no grade is named", async () => {
        alexIsSaved();
        await runInvite(deps({ chooseContact: picksAlex }));
        expect(runtime.invited).toEqual([{ endpointId: ALEX, role: "editor" }]);
    });

    // The corner is the only confirmation of what was granted, so it has to
    // say which of the two things happened.
    it("names the grade it granted, and the partner it granted it to", async () => {
        alexIsSaved();
        const viewing = deps({ chooseContact: picksAlex });
        await runInvite(viewing, "viewer");
        expect(viewing.notices).toEqual(["Invited Alex to view"]);

        const editing = deps({ chooseContact: picksAlex });
        await runInvite(editing, "editor");
        expect(editing.notices).toEqual(["Invited Alex to edit"]);
    });
});
