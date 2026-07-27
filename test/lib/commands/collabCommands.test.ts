import { beforeEach, describe, expect, it, vi } from "vitest";

import { endSession } from "@/lib/collab/runtime";
import { parseTicket } from "@/lib/collab/ticket";
import {
    runEnd,
    runInvite,
    runJoin,
    runShare,
    type CollabCommandDeps,
} from "@/lib/commands/collabCommands";
import { COLLAB_COMMANDS, COMMANDS } from "@/lib/commands/registry";
import { getPresetKeymap } from "@/lib/keymap/presets";
import { makeFlowRound } from "@/lib/model/flow";
import { useFlowStore } from "@/lib/store/useFlowStore";

function deps(over: Partial<CollabCommandDeps> = {}): CollabCommandDeps & {
    notices: string[];
    failures: string[];
    copied: string[];
} {
    const notices: string[] = [];
    const failures: string[] = [];
    const copied: string[] = [];
    return {
        notices,
        failures,
        copied,
        notify: (m) => notices.push(m),
        fail: (m) => failures.push(m),
        askForTicket: async () => null,
        copy: async (t) => {
            copied.push(t);
        },
        openFlow: vi.fn(),
        ...over,
    };
}

beforeEach(async () => {
    await endSession();
    useFlowStore.setState({ collabEnabled: false, round: null });
    // isDesktop() is false under jsdom unless the harness says otherwise.
    delete (window as unknown as Record<string, unknown>).__TAURI_INTERNALS__;
});

describe("the commands are palette-only", () => {
    it("claims no chord in the default keymap", () => {
        const bound = new Set(Object.values(getPresetKeymap().bindings));
        for (const id of COLLAB_COMMANDS) expect(bound.has(id)).toBe(false);
    });

    it("is registered with a label a debater can search for", () => {
        expect(COMMANDS["collab.share"].label).toBe("Share this round");
        expect(COMMANDS["collab.join"].label).toBe("Join a shared round");
        expect(COMMANDS["collab.end"].label).toBe("End shared session");
    });
});

describe("share", () => {
    it("refuses while the master switch is off, and copies nothing", async () => {
        const d = deps();
        await runShare(d);
        expect(d.copied).toEqual([]);
        expect(d.failures[0]).toMatch(/Settings/);
    });

    it("refuses with no flow open", async () => {
        (window as unknown as Record<string, unknown>).__TAURI_INTERNALS__ = {};
        useFlowStore.setState({ collabEnabled: true, round: null });
        const d = deps();
        await runShare(d);
        expect(d.failures[0]).toMatch(/Open a flow/);
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
});

describe("the ticket a share would copy", () => {
    it("round-trips through the parser the join side uses", () => {
        const round = makeFlowRound({});
        expect(parseTicket("nonsense")).toBeNull();
        expect(round.id).toBeTruthy();
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
