import { beforeEach, describe, expect, it, vi } from "vitest";

import type { ShadowEntry } from "@/lib/collab/shadow";
import { type CollabPeerView, useCollabStore } from "@/lib/store/useCollabStore";
import { useFlowStore } from "@/lib/store/useFlowStore";

const ALEX: CollabPeerView = {
    endpointId: "alex",
    name: "Alex",
    role: "partner",
    connectionType: "direct",
};

const COACH: CollabPeerView = {
    endpointId: "rin",
    name: "Rin",
    role: "coach",
    connectionType: "relayed",
};

function observation(from: string, at: number): ShadowEntry {
    return { at, from, diffs: [], dropped: [] };
}

beforeEach(() => {
    useCollabStore.getState().reset();
    useCollabStore.getState().clearShadow();
});

describe("useCollabStore", () => {
    it("starts off with no peers", () => {
        expect(useCollabStore.getState().status).toBe("off");
        expect(useCollabStore.getState().peers).toEqual([]);
    });

    it("sets the connection state", () => {
        useCollabStore.getState().setStatus("connecting");
        expect(useCollabStore.getState().status).toBe("connecting");

        useCollabStore.getState().setStatus("connected");
        expect(useCollabStore.getState().status).toBe("connected");

        useCollabStore.getState().setStatus("reconnecting");
        expect(useCollabStore.getState().status).toBe("reconnecting");
    });

    it("sets the peer list", () => {
        useCollabStore.getState().setPeers([ALEX, COACH]);
        expect(useCollabStore.getState().peers).toEqual([ALEX, COACH]);
    });

    it("reset clears a populated session back to the defaults", () => {
        useCollabStore.getState().setStatus("connected");
        useCollabStore.getState().setPeers([ALEX, COACH]);

        useCollabStore.getState().reset();

        expect(useCollabStore.getState().status).toBe("off");
        expect(useCollabStore.getState().peers).toEqual([]);
    });

    it("notifies no flow-store subscriber, so presence never re-renders the grid", () => {
        const seen = vi.fn();
        const unsubscribe = useFlowStore.subscribe(seen);

        useCollabStore.getState().setStatus("connected");
        useCollabStore.getState().setPeers([ALEX]);
        unsubscribe();

        expect(seen).not.toHaveBeenCalled();
    });
});

describe("useCollabStore shadow log", () => {
    it("starts empty", () => {
        expect(useCollabStore.getState().shadowLog).toEqual([]);
    });

    it("keeps observations in the order they arrived", () => {
        useCollabStore.getState().pushShadow(observation("alex", 1));
        useCollabStore.getState().pushShadow(observation("rin", 2));

        expect(useCollabStore.getState().shadowLog.map((e) => e.from)).toEqual(["alex", "rin"]);
    });

    it("drops the oldest once a long round fills the log", () => {
        for (let at = 1; at <= 260; at++) {
            useCollabStore.getState().pushShadow(observation("alex", at));
        }
        const log = useCollabStore.getState().shadowLog;

        expect(log).toHaveLength(200);
        expect(log[0].at).toBe(61);
        expect(log[log.length - 1].at).toBe(260);
    });

    it("clears on request", () => {
        useCollabStore.getState().pushShadow(observation("alex", 1));
        useCollabStore.getState().clearShadow();

        expect(useCollabStore.getState().shadowLog).toEqual([]);
    });

    it("survives reset, so a finished session is still readable", () => {
        useCollabStore.getState().setStatus("connected");
        useCollabStore.getState().pushShadow(observation("alex", 1));

        useCollabStore.getState().reset();

        expect(useCollabStore.getState().status).toBe("off");
        expect(useCollabStore.getState().shadowLog).toHaveLength(1);
    });

    it("notifies no flow-store subscriber, so a recorded diff never re-renders the grid", () => {
        const seen = vi.fn();
        const unsubscribe = useFlowStore.subscribe(seen);

        useCollabStore.getState().pushShadow(observation("alex", 1));
        useCollabStore.getState().clearShadow();
        unsubscribe();

        expect(seen).not.toHaveBeenCalled();
    });
});

describe("useCollabStore invitations", () => {
    const invite = { endpointId: "alex", roundId: "r1", label: "Round 3" };

    beforeEach(() => {
        useCollabStore.getState().dismissInvite("alex", "r1");
        useCollabStore.getState().dismissInvite("alex", "r2");
    });

    it("starts with none", () => {
        expect(useCollabStore.getState().invites).toEqual([]);
    });

    it("keeps one offer per round, however many times a partner dials", () => {
        useCollabStore.getState().pushInvite(invite);
        useCollabStore.getState().pushInvite({ ...invite, label: "Round 3 again" });
        expect(useCollabStore.getState().invites).toEqual([invite]);
    });

    it("keeps a second round from the same partner apart from the first", () => {
        useCollabStore.getState().pushInvite(invite);
        useCollabStore.getState().pushInvite({ ...invite, roundId: "r2" });
        expect(useCollabStore.getState().invites.map((i) => i.roundId)).toEqual(["r1", "r2"]);
    });

    it("drops exactly the one that was acted on", () => {
        useCollabStore.getState().pushInvite(invite);
        useCollabStore.getState().pushInvite({ ...invite, roundId: "r2" });
        useCollabStore.getState().dismissInvite("alex", "r1");
        expect(useCollabStore.getState().invites.map((i) => i.roundId)).toEqual(["r2"]);
    });

    it("outlives the session that carried it, which may end before anyone looks", () => {
        useCollabStore.getState().pushInvite(invite);
        useCollabStore.getState().reset();
        expect(useCollabStore.getState().invites).toEqual([invite]);
    });
});
