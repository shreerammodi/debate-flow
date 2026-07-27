import { beforeEach, describe, expect, it, vi } from "vitest";

interface Corner {
    message: string;
    id?: string;
    action?: { label: string; onClick: () => void };
}

const corners: Corner[] = [];
const errors: string[] = [];

/** Sonner keeps one toast per id, so a repeat replaces rather than stacks. */
vi.mock("sonner", () => ({
    toast: Object.assign(
        (
            message: string,
            opts?: { id?: string; action?: { label: string; onClick: () => void } },
        ) => {
            const corner = { message, action: opts?.action };
            const at = opts?.id ? corners.findIndex((c) => c.id === opts.id) : -1;
            if (at >= 0) corners[at] = { ...corner, id: opts.id };
            else corners.push({ ...corner, id: opts?.id });
        },
        {
            warning: () => {},
            error: (m: string) => errors.push(m),
            success: () => {},
            info: () => {},
        },
    ),
}));

const routed: string[] = [];

vi.mock("@/lib/commands/flowNav", async (importOriginal) => ({
    ...(await importOriginal<typeof import("@/lib/commands/flowNav")>()),
    navigateToFlow: (path: string) => routed.push(path),
}));

const joins: { calls: unknown[]; result: unknown; fail: Error | null } = {
    calls: [],
    result: null,
    fail: null,
};

vi.mock("@/lib/collab/join", () => ({
    joinRound: (deps: unknown) => {
        joins.calls.push(deps);
        if (joins.fail) return Promise.reject(joins.fail);
        return Promise.resolve(joins.result);
    },
}));

import { acceptInvite, announceInvite } from "@/lib/collab/inbox";
import { useCollabStore } from "@/lib/store/useCollabStore";
import { useFlowStore } from "@/lib/store/useFlowStore";

const ALEX = "alex";
const notice = { endpointId: ALEX, roundId: "r1", label: "Round 3 - Harvard" };

beforeEach(() => {
    corners.length = 0;
    errors.length = 0;
    routed.length = 0;
    joins.calls.length = 0;
    joins.fail = null;
    joins.result = { roundId: "r1", hostEndpointId: ALEX, path: "/flows/r3.ebb", created: true };
    useCollabStore.getState().dismissInvite(ALEX, "r1");
    useFlowStore.setState({ contacts: { [ALEX]: { name: "Alex", role: "partner" } } });
});

describe("announceInvite", () => {
    it("names the partner and the round in the corner", () => {
        announceInvite(notice);
        expect(corners[0].message).toBe("Alex shared Round 3 - Harvard");
        expect(corners[0].action!.label).toBe("Join");
    });

    it("holds the offer, so the start screen can show it too", () => {
        announceInvite(notice);
        expect(useCollabStore.getState().invites).toEqual([notice]);
    });

    it("says nothing at all for a peer nobody saved", () => {
        useFlowStore.setState({ contacts: {} });
        announceInvite(notice);
        expect(corners).toEqual([]);
        expect(useCollabStore.getState().invites).toEqual([]);
    });

    // One share can reach this machine as two dials, and a contact may offer
    // the same round again after the first message has gone.
    it("refreshes one message when the same round is offered twice", () => {
        announceInvite(notice);
        announceInvite(notice);
        expect(corners).toHaveLength(1);
        expect(corners[0].message).toBe("Alex shared Round 3 - Harvard");
        expect(useCollabStore.getState().invites).toEqual([notice]);
    });

    it("keeps a second round from the same partner as its own message", () => {
        announceInvite(notice);
        announceInvite({ ...notice, roundId: "r2", label: "Round 4 - Bronx" });
        expect(corners.map((c) => c.message)).toEqual([
            "Alex shared Round 3 - Harvard",
            "Alex shared Round 4 - Bronx",
        ]);
        useCollabStore.getState().dismissInvite(ALEX, "r2");
    });

    it("joins nothing on its own", () => {
        announceInvite(notice);
        expect(joins.calls).toEqual([]);
        expect(routed).toEqual([]);
    });
});

describe("acceptInvite", () => {
    it("asks the host for the round by EndpointId, with no ticket", async () => {
        await acceptInvite(notice);
        expect(joins.calls).toHaveLength(1);
        const deps = joins.calls[0] as { ticket?: string; invite?: unknown };
        expect(deps.ticket).toBeUndefined();
        expect(deps.invite).toEqual({ endpointId: ALEX, roundId: "r1" });
    });

    it("opens the file the join landed in", async () => {
        await acceptInvite(notice);
        expect(routed).toEqual(["/flows/r3.ebb"]);
    });

    it("clears the offer once it has been taken", async () => {
        announceInvite(notice);
        await acceptInvite(notice);
        expect(useCollabStore.getState().invites).toEqual([]);
    });

    it("keeps the offer when the join fails, so it can be tried again", async () => {
        announceInvite(notice);
        joins.fail = new Error("The host hung up");
        await acceptInvite(notice);
        expect(errors).toEqual(["The host hung up"]);
        expect(useCollabStore.getState().invites).toEqual([notice]);
        expect(routed).toEqual([]);
    });

    it("says what to switch on when shared editing is off", async () => {
        joins.result = null;
        await acceptInvite(notice);
        expect(errors).toEqual(["Turn on shared editing in Settings first"]);
    });
});
