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

function projected(base: FlowRound): FlowRound {
    return projectDoc(getReplica()!, base);
}

describe("structural store actions reach the replica", () => {
    it("mirrors a sheet added, renamed, removed, and restored", () => {
        const round = makeFlowRound({});
        const store = useFlowStore.getState();
        store.loadRound(round);

        const id = store.addSheet({ group: "neg" });
        expect(projected(round).sheets.map((s) => s.id)).toContain(id);

        store.renameSheet(id, "Topicality");
        expect(projected(round).sheets.find((s) => s.id === id)!.title).toBe("Topicality");

        const removed = store.removeSheet(id);
        expect(projected(round).sheets.map((s) => s.id)).not.toContain(id);

        store.restoreSheet(removed!);
        expect(projected(round).sheets.map((s) => s.id)).toContain(id);
    });

    it("mirrors every sheet a batch add creates", () => {
        const round = makeFlowRound({});
        const store = useFlowStore.getState();
        store.loadRound(round);
        const ids = store.addSheets([{ group: "aff" }, { group: "neg" }]);
        const seen = projected(round).sheets.map((s) => s.id);
        for (const id of ids) expect(seen).toContain(id);
    });

    it("mirrors a reorder across every sheet it renumbers", () => {
        const round = makeFlowRound({});
        const store = useFlowStore.getState();
        store.loadRound(round);
        const a = store.addSheet({ group: "aff" });
        const b = store.addSheet({ group: "neg" });
        store.reorderSheets([b, a]);
        const orders = new Map(projected(round).sheets.map((s) => [s.id, s.order] as const));
        expect(orders.get(b)!).toBeLessThan(orders.get(a)!);
    });

    it("mirrors a scouting edit, including a nested decision leaf", () => {
        const round = makeFlowRound({});
        const store = useFlowStore.getState();
        store.loadRound(round);
        store.setScouting({ tournament: "Harvard", judge: "Ito" });
        store.setScouting({ decision: { vote: "aff", rfd: "turns outweigh" } });
        const s = projected(round).scouting;
        expect(s.tournament).toBe("Harvard");
        expect(s.judge).toBe("Ito");
        expect(s.decision).toEqual({ vote: "aff", rfd: "turns outweigh" });
    });

    it("mirrors a speaking-order swap", () => {
        const round = makeFlowRound({ event: "pf", firstSide: "aff" });
        const store = useFlowStore.getState();
        store.loadRound(round);
        store.swapSpeakingOrder();
        expect(projected(round).firstSide).toBe("neg");
    });

    it("records nothing when an action no-ops", () => {
        const round = makeFlowRound({ event: "policy" });
        const store = useFlowStore.getState();
        store.loadRound(round);
        const before = getReplica();
        // Policy has a fixed speaking order, so the swap returns early.
        store.swapSpeakingOrder();
        expect(getReplica()).toBe(before);
    });
});
