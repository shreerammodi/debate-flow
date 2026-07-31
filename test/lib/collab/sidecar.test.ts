import { describe, expect, it } from "vitest";

import { seedDoc } from "@/lib/collab/doc";
import { parseSidecar, serializeSidecar, SIDECAR_VERSION } from "@/lib/collab/sidecar";
import { makeFlowRound } from "@/lib/model/flow";

const round = makeFlowRound({});
const doc = seedDoc(round);
const SAM = "5".repeat(64);
const KIM = "k".repeat(52);
const text = serializeSidecar({
    roundId: round.id,
    flowHash: "deadbeef",
    peers: [SAM, KIM],
    coaches: [KIM],
    relays: {},
    doc,
});

describe("serializeSidecar", () => {
    it("stamps the version, the round, and the file it belongs to", () => {
        const parsed = JSON.parse(text);
        expect(parsed.version).toBe(SIDECAR_VERSION);
        expect(parsed.roundId).toBe(round.id);
        expect(parsed.flowHash).toBe("deadbeef");
        expect(parsed.peers).toEqual([SAM, KIM]);
        expect(parsed.coaches).toEqual([KIM]);
    });
});

describe("parseSidecar", () => {
    it("recovers a matching sidecar", () => {
        const got = parseSidecar(text, round.id, "deadbeef");
        expect(got!.doc).toEqual(doc);
        expect(got!.peers).toEqual([SAM, KIM]);
        expect(got!.coaches).toEqual([KIM]);
    });

    // Every entry here is dialled on the next open, so a hand edit or a peer's
    // junk must not survive the read.
    it("keeps only ids iroh could parse back into a key", () => {
        const edited = JSON.parse(text);
        edited.peers = [SAM, "sam", 7, null, "../../etc/passwd"];
        edited.coaches = ["nope"];
        const got = parseSidecar(JSON.stringify(edited), round.id, "deadbeef")!;
        expect(got.peers).toEqual([SAM]);
        expect(got.coaches).toEqual([]);
    });

    /**
     * A relay is a dial target the next time this round opens, so a scheme
     * somebody chose by hand is not one, and neither is a string too long to
     * be an address. Addressing and not admission, so junk here drops the
     * address and never the round.
     */
    it("keeps only the https relays, against the peers they name", () => {
        const edited = JSON.parse(text);
        edited.relays = {
            [SAM]: "https://usw1-1.relay.n0.iroh.link./",
            [KIM]: "http://relay.example/",
            ["7".repeat(64)]: `https://relay.example/${"a".repeat(256)}`,
            nobody: "https://relay.example/",
            __proto__: "https://relay.example/",
        };
        const got = parseSidecar(JSON.stringify(edited), round.id, "deadbeef")!;
        expect(got.relays).toEqual({ [SAM]: "https://usw1-1.relay.n0.iroh.link./" });
    });

    // Written by a build that recorded no addresses, which is a round whose
    // peers are reachable across the room and no further - exactly what that
    // build could do. Not a reason to throw the document away.
    it("reads a sidecar with no relays at all as a round with none", () => {
        const noRelays = JSON.parse(text);
        delete noRelays.relays;
        expect(parseSidecar(JSON.stringify(noRelays), round.id, "deadbeef")!.relays).toEqual({});
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

    // A file written before `coaches` existed holds membership and no grades,
    // so every peer it remembers reads back a partner: one silent promotion
    // per already-shared round, on the first open after an upgrade. Discarding
    // it costs a re-seed of the replica, which is what a stale one costs too.
    it("discards a sidecar written before read-only grants were recorded", () => {
        const before = JSON.stringify({
            version: 1,
            roundId: round.id,
            flowHash: "deadbeef",
            peers: [SAM, KIM],
            doc,
        });
        expect(parseSidecar(before, round.id, "deadbeef")).toBeNull();
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
