import { describe, expect, it } from "vitest";

import { seedDoc } from "@/lib/collab/doc";
import { parseSidecar, serializeSidecar, SIDECAR_VERSION } from "@/lib/collab/sidecar";
import { makeFlowRound } from "@/lib/model/flow";

const round = makeFlowRound({});
const doc = seedDoc(round);
const text = serializeSidecar({
    roundId: round.id,
    flowHash: "deadbeef",
    peers: ["sam"],
    doc,
});

describe("serializeSidecar", () => {
    it("stamps the version, the round, and the file it belongs to", () => {
        const parsed = JSON.parse(text);
        expect(parsed.version).toBe(SIDECAR_VERSION);
        expect(parsed.roundId).toBe(round.id);
        expect(parsed.flowHash).toBe("deadbeef");
        expect(parsed.peers).toEqual(["sam"]);
    });
});

describe("parseSidecar", () => {
    it("recovers a matching sidecar", () => {
        const got = parseSidecar(text, round.id, "deadbeef");
        expect(got!.doc).toEqual(doc);
        expect(got!.peers).toEqual(["sam"]);
    });

    it("discards a sidecar whose flow has moved on", () => {
        expect(parseSidecar(text, round.id, "cafe1234")).toBeNull();
    });

    it("discards a sidecar for another round", () => {
        expect(parseSidecar(text, "round_other", "deadbeef")).toBeNull();
    });

    it("discards a version it does not know", () => {
        const future = JSON.stringify({ ...JSON.parse(text), version: 99 });
        expect(parseSidecar(future, round.id, "deadbeef")).toBeNull();
    });

    it("discards an absent, empty, or malformed file rather than throwing", () => {
        expect(parseSidecar(null, round.id, "deadbeef")).toBeNull();
        expect(parseSidecar("", round.id, "deadbeef")).toBeNull();
        expect(parseSidecar("{not json", round.id, "deadbeef")).toBeNull();
        expect(parseSidecar("[]", round.id, "deadbeef")).toBeNull();
        expect(parseSidecar("null", round.id, "deadbeef")).toBeNull();
    });

    it("discards one whose document is the wrong shape", () => {
        const broken = JSON.stringify({ ...JSON.parse(text), doc: { roundId: round.id } });
        expect(parseSidecar(broken, round.id, "deadbeef")).toBeNull();
    });

    it("defaults a missing peer list rather than discarding the document", () => {
        const noPeers = JSON.parse(text);
        delete noPeers.peers;
        expect(parseSidecar(JSON.stringify(noPeers), round.id, "deadbeef")!.peers).toEqual([]);
    });
});
