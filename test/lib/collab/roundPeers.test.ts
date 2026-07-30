import { beforeEach, describe, expect, it } from "vitest";

import {
    forgetRoundPeer,
    forgetRoundPeers,
    knownRoundCoaches,
    knownRoundPeers,
    rememberRoundPeers,
    rememberRoundRole,
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
        setRoundPeers("r1", ["sam", "kim"], []);
        expect(knownRoundPeers("r1")).toEqual(["sam", "kim"]);
    });

    it("keeps a peer who is not connected right now", () => {
        setRoundPeers("r1", ["sam", "kim"], []);
        rememberRoundPeers("r1", ["sam"]);
        expect(knownRoundPeers("r1")).toEqual(["sam", "kim"]);
    });

    it("counts one peer once, however many times they connect", () => {
        setRoundPeers("r1", ["sam"], []);
        rememberRoundPeers("r1", ["sam", "sam", "kim"]);
        expect(knownRoundPeers("r1")).toEqual(["sam", "kim"]);
    });

    it("answers for a round it was told about and no other", () => {
        setRoundPeers("r1", ["sam"], []);
        expect(knownRoundPeers("r2")).toEqual([]);
    });

    // A join records its host under the round being joined, which is never the
    // round on screen. One slot for all of them let that join spend the open
    // round's membership, and the open round's next autosave wrote the loss out.
    it("keeps one round's set when another round is remembered", () => {
        setRoundPeers("r1", ["sam"], []);
        rememberRoundPeers("r2", ["kim"]);
        expect(knownRoundPeers("r2")).toEqual(["kim"]);
        expect(knownRoundPeers("r1")).toEqual(["sam"]);
    });

    it("keeps one round's read-only grades when another round is remembered", () => {
        setRoundPeers("r1", ["kim"], ["kim"]);
        rememberRoundPeers("r2", ["sam"]);
        expect(knownRoundCoaches("r1")).toEqual(["kim"]);
    });
});

describe("cutting one peer out of what a round remembers", () => {
    // The set is otherwise append-only, so this is the whole of revocation:
    // without it the next open re-dials the peer off the sidecar and admits
    // them on membership alone.
    it("drops the peer and keeps the rest", () => {
        setRoundPeers("r1", ["sam", "kim"], []);
        forgetRoundPeer("r1", "sam");
        expect(knownRoundPeers("r1")).toEqual(["kim"]);
    });

    it("drops their read-only mark with them", () => {
        setRoundPeers("r1", ["sam", "kim"], ["sam"]);
        forgetRoundPeer("r1", "sam");
        expect(knownRoundCoaches("r1")).toEqual([]);
    });

    it("leaves another round's set alone", () => {
        setRoundPeers("r1", ["sam"], []);
        setRoundPeers("r2", ["sam"], []);
        forgetRoundPeer("r2", "sam");
        expect(knownRoundPeers("r1")).toEqual(["sam"]);
    });

    it("says nothing about a peer who was never there", () => {
        setRoundPeers("r1", ["sam"], []);
        forgetRoundPeer("r1", "kim");
        expect(knownRoundPeers("r1")).toEqual(["sam"]);
    });
});

describe("what a round remembers a peer was admitted as", () => {
    it("grades nobody until somebody says so", () => {
        setRoundPeers("r1", ["sam"], []);
        expect(knownRoundCoaches("r1")).toEqual([]);
    });

    it("takes the marks a sidecar recovered", () => {
        setRoundPeers("r1", ["sam", "kim"], ["kim"]);
        expect(knownRoundCoaches("r1")).toEqual(["kim"]);
    });

    // A grant the contact table never saw has to live somewhere durable, or
    // the round remembers the membership and forgets the restriction.
    it("marks a peer admitted read-only, and remembers them as a peer", () => {
        setRoundPeers("r1", [], []);
        rememberRoundRole("r1", "kim", "coach");
        expect(knownRoundCoaches("r1")).toEqual(["kim"]);
        expect(knownRoundPeers("r1")).toEqual(["kim"]);
    });

    it("clears the mark when the same peer is admitted wider", () => {
        setRoundPeers("r1", ["kim"], ["kim"]);
        rememberRoundRole("r1", "kim", "partner");
        expect(knownRoundCoaches("r1")).toEqual([]);
        expect(knownRoundPeers("r1")).toEqual(["kim"]);
    });

    it("counts one mark once, however many times the peer reconnects", () => {
        setRoundPeers("r1", ["kim"], []);
        rememberRoundRole("r1", "kim", "coach");
        rememberRoundRole("r1", "kim", "coach");
        expect(knownRoundCoaches("r1")).toEqual(["kim"]);
    });

    // The round a session grades is its own, which after a join is not the
    // round this module was last asked about. Dropping the grade there is the
    // same promotion by a quieter route.
    it("grades a round that is not the one last remembered", () => {
        setRoundPeers("r1", ["kim"], []);
        rememberRoundPeers("r2", ["sam"]);
        rememberRoundRole("r1", "kim", "coach");
        expect(knownRoundCoaches("r1")).toEqual(["kim"]);
        expect(knownRoundCoaches("r2")).toEqual([]);
    });

    it("answers for the round it was told about and no other", () => {
        setRoundPeers("r1", ["kim"], ["kim"]);
        expect(knownRoundCoaches("r2")).toEqual([]);
    });

    // Back at the start screen nothing is open, so holding any round's partner
    // ids costs privacy for nothing.
    it("forgets every round's marks when the last round closes", () => {
        setRoundPeers("r1", ["kim"], ["kim"]);
        setRoundPeers("r2", ["sam"], ["sam"]);
        forgetRoundPeers();
        expect(knownRoundCoaches("r1")).toEqual([]);
        expect(knownRoundPeers("r2")).toEqual([]);
    });

    // Remembering is what the runtime does on every peer change, and it must
    // not disturb a grade the session already recorded.
    it("keeps a mark across a plain remember", () => {
        setRoundPeers("r1", ["kim"], ["kim"]);
        rememberRoundPeers("r1", ["kim", "sam"]);
        expect(knownRoundCoaches("r1")).toEqual(["kim"]);
        expect(knownRoundPeers("r1")).toEqual(["kim", "sam"]);
    });
});
