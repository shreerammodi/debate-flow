import { beforeEach, describe, expect, it, vi } from "vitest";

import { type CollabPeerView, useCollabStore } from "@/lib/store/useCollabStore";
import { useFlowStore } from "@/lib/store/useFlowStore";

const ALEX: CollabPeerView = {
    endpointId: "alex",
    name: "Alex",
    role: "editor",
    connectionType: "direct",
};

const COACH: CollabPeerView = {
    endpointId: "rin",
    name: "Rin",
    role: "viewer",
    connectionType: "relayed",
};

beforeEach(() => {
    useCollabStore.getState().reset();
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
        useCollabStore.getState().pushInvite({ endpointId: "alex", roundId: "r1", label: "" });

        useCollabStore.getState().reset();

        expect(useCollabStore.getState().status).toBe("off");
        expect(useCollabStore.getState().peers).toEqual([]);
        expect(useCollabStore.getState().invites).toEqual([]);
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

describe("useCollabStore invitations", () => {
    const invite = { endpointId: "alex", roundId: "r1", label: "Round 3" };

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

    it("holds nothing across a teardown, because a dead session can join none of it", () => {
        useCollabStore.getState().pushInvite(invite);
        useCollabStore.getState().reset();
        expect(useCollabStore.getState().invites).toEqual([]);
    });

    it("bounds the list, so a partner minting round ids cannot grow it forever", () => {
        for (let i = 0; i < 500; i++) {
            useCollabStore.getState().pushInvite({ ...invite, roundId: `r${i}` });
        }
        const held = useCollabStore.getState().invites;
        expect(held.length).toBeLessThanOrEqual(20);
        // The newest survive: their sender is the one still holding a round up.
        expect(held[held.length - 1].roundId).toBe("r499");
        expect(held.some((i) => i.roundId === "r0")).toBe(false);
    });
});
