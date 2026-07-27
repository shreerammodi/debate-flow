import { beforeEach, describe, expect, it } from "vitest";

import { projectDoc } from "@/lib/collab/doc";
import { getReplica, replicaRoundId } from "@/lib/collab/replica";
import { makeFlowRound, type FlowRound } from "@/lib/model/flow";
import { useFlowStore } from "@/lib/store/useFlowStore";

beforeEach(() => {
    useFlowStore.getState().closeRound();
});

describe("the replica follows the open round", () => {
    it("seeds when a round opens", () => {
        const round = makeFlowRound({});
        useFlowStore.getState().loadRound(round);
        expect(replicaRoundId()).toBe(round.id);
        expect(Object.keys(getReplica()!.sheets)).toHaveLength(round.sheets.length);
    });

    it("re-seeds when a second flow opens straight over the first", () => {
        const first = makeFlowRound({});
        const second = makeFlowRound({});
        useFlowStore.getState().loadRound(first);
        useFlowStore.getState().loadRound(second);
        expect(replicaRoundId()).toBe(second.id);
    });

    it("drops the replica when the round closes", () => {
        useFlowStore.getState().loadRound(makeFlowRound({}));
        useFlowStore.getState().closeRound();
        expect(getReplica()).toBeNull();
    });
});
