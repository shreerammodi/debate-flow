import { beforeEach, describe, expect, it, vi } from "vitest";

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
