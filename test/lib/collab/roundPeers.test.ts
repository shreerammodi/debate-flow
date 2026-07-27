import { beforeEach, describe, expect, it } from "vitest";

import {
    forgetRoundPeers,
    knownRoundPeers,
    rememberRoundPeers,
    setRoundPeers,
} from "@/lib/collab/roundPeers";

beforeEach(() => {
    forgetRoundPeers();
});

describe("the peers a round remembers", () => {
    it("knows nobody for a round that was never shared", () => {
        expect(knownRoundPeers("r1")).toEqual([]);
    });

    it("takes the set a sidecar recovered", () => {
        setRoundPeers("r1", ["sam", "kim"]);
        expect(knownRoundPeers("r1")).toEqual(["sam", "kim"]);
    });

    it("keeps a peer who is not connected right now", () => {
        setRoundPeers("r1", ["sam", "kim"]);
        rememberRoundPeers("r1", ["sam"]);
        expect(knownRoundPeers("r1")).toEqual(["sam", "kim"]);
    });

    it("counts one peer once, however many times they connect", () => {
        setRoundPeers("r1", ["sam"]);
        rememberRoundPeers("r1", ["sam", "sam", "kim"]);
        expect(knownRoundPeers("r1")).toEqual(["sam", "kim"]);
    });

    it("answers for the tracked round and no other", () => {
        setRoundPeers("r1", ["sam"]);
        expect(knownRoundPeers("r2")).toEqual([]);
    });

    it("starts over when a different round is opened", () => {
        setRoundPeers("r1", ["sam"]);
        rememberRoundPeers("r2", ["kim"]);
        expect(knownRoundPeers("r2")).toEqual(["kim"]);
        expect(knownRoundPeers("r1")).toEqual([]);
    });
});
