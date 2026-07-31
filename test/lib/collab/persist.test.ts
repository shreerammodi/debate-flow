import { beforeEach, describe, expect, it } from "vitest";

import { seedDoc } from "@/lib/collab/doc";
import { hashText } from "@/lib/collab/hash";
import { persistReplica, recoverReplica } from "@/lib/collab/persist";
import { clearReplica, getReplica } from "@/lib/collab/replica";
import { parseSidecar, serializeSidecar } from "@/lib/collab/sidecar";
import { setSidecarFs, type SidecarFs } from "@/lib/collab/sidecarFs";
import { makeFlowRound, type FlowRound } from "@/lib/model/flow";
import { serializeFlow } from "@/lib/persistence/flowFile";
import { useFlowStore } from "@/lib/store/useFlowStore";

interface FakeSidecarFs extends SidecarFs {
    files: Map<string, string>;
}

function fakeFs(): FakeSidecarFs {
    const files = new Map<string, string>();
    return {
        files,
        async read(id) {
            return files.get(id) ?? null;
        },
        async write(id, text) {
            files.set(id, text);
        },
    };
}

let fs: FakeSidecarFs;
let round: FlowRound;
/** Real-shaped ids: the sidecar drops anything that is not one, on the way in. */
const SAM = "5".repeat(64);
const KIM = "c".repeat(64);

beforeEach(() => {
    fs = fakeFs();
    setSidecarFs(fs);
    clearReplica();
    useFlowStore.setState({ collabEnabled: true });
    round = makeFlowRound({});
    round.sheets.find((s) => s.kind !== "cx")!.data = [["perm", "link"]];
});

describe("recoverReplica", () => {
    it("seeds from the file when no sidecar exists", async () => {
        await recoverReplica(round, serializeFlow(round));
        expect(getReplica()).toEqual(seedDoc(round));
    });

    it("recovers a sidecar whose hash still names the file", async () => {
        const text = serializeFlow(round);
        const doc = seedDoc(round);
        doc.round.event = { value: "pf", stamp: { ms: 7, counter: 0, actor: "sam" } };
        fs.files.set(
            round.id,
            serializeSidecar({
                roundId: round.id,
                flowHash: hashText(text),
                peers: [],
                coaches: [],
                relays: {},
                doc,
            }),
        );
        await recoverReplica(round, text);
        expect(getReplica()!.round.event.value).toBe("pf");
    });

    it("discards a sidecar written against an older file", async () => {
        const doc = seedDoc(round);
        doc.round.event = { value: "pf", stamp: { ms: 7, counter: 0, actor: "sam" } };
        fs.files.set(
            round.id,
            serializeSidecar({
                roundId: round.id,
                flowHash: "stale000",
                peers: [],
                coaches: [],
                relays: {},
                doc,
            }),
        );
        await recoverReplica(round, serializeFlow(round));
        expect(getReplica()!.round.event.value).toBe("policy");
    });

    it("still seeds from the file when shared editing is off", async () => {
        useFlowStore.setState({ collabEnabled: false });
        await recoverReplica(round, serializeFlow(round));
        expect(getReplica()).toEqual(seedDoc(round));
        expect(fs.files.size).toBe(0);
    });

    it("seeds from the file when the sidecar cannot be read at all", async () => {
        setSidecarFs({
            read: async () => {
                throw new Error("no config directory");
            },
            write: async () => {},
        });
        await recoverReplica(round, serializeFlow(round));
        expect(getReplica()).toEqual(seedDoc(round));
    });

    it("reports the peers the round was shared with", async () => {
        const text = serializeFlow(round);
        fs.files.set(
            round.id,
            serializeSidecar({
                roundId: round.id,
                flowHash: hashText(text),
                peers: [SAM, KIM],
                coaches: [],
                relays: {},
                doc: seedDoc(round),
            }),
        );
        expect(await recoverReplica(round, text)).toEqual([SAM, KIM]);
    });

    it("reports nobody for a round that was never shared", async () => {
        expect(await recoverReplica(round, serializeFlow(round))).toEqual([]);
    });

    it("reports nobody at all when shared editing is off", async () => {
        const text = serializeFlow(round);
        fs.files.set(
            round.id,
            serializeSidecar({
                roundId: round.id,
                flowHash: hashText(text),
                peers: [SAM],
                coaches: [],
                relays: {},
                doc: seedDoc(round),
            }),
        );
        useFlowStore.setState({ collabEnabled: false });
        expect(await recoverReplica(round, text)).toEqual([]);
    });
});

describe("persistReplica", () => {
    it("writes a sidecar stamped with the file it belongs to", async () => {
        const text = serializeFlow(round);
        await recoverReplica(round, text);
        await persistReplica(round, text);
        // Read back through the real gate: a sidecar this build cannot
        // recover is not a sidecar worth having written.
        const written = parseSidecar(fs.files.get(round.id)!, round.id, hashText(text));
        expect(written).not.toBeNull();
        expect(written!.roundId).toBe(round.id);
    });

    it("carries the round's peers forward, so the next open re-dials them", async () => {
        const text = serializeFlow(round);
        fs.files.set(
            round.id,
            serializeSidecar({
                roundId: round.id,
                flowHash: hashText(text),
                peers: [SAM],
                coaches: [],
                relays: {},
                doc: seedDoc(round),
            }),
        );
        await recoverReplica(round, text);
        await persistReplica(round, text);
        const written = parseSidecar(fs.files.get(round.id)!, round.id, hashText(text));
        expect(written!.peers).toEqual([SAM]);
    });

    // The grant lived only in the contact table, so removing a coach there
    // promoted them to partner the next time the round opened.
    it("carries a read-only grant forward, so a coach is not promoted on the next open", async () => {
        const text = serializeFlow(round);
        fs.files.set(
            round.id,
            serializeSidecar({
                roundId: round.id,
                flowHash: hashText(text),
                peers: [SAM, KIM],
                coaches: [KIM],
                relays: {},
                doc: seedDoc(round),
            }),
        );
        await recoverReplica(round, text);
        await persistReplica(round, text);
        const written = parseSidecar(fs.files.get(round.id)!, round.id, hashText(text));
        expect(written!.peers).toEqual([SAM, KIM]);
        expect(written!.coaches).toEqual([KIM]);
    });

    it("heals a drifted sheet before it writes", async () => {
        await recoverReplica(round, serializeFlow(round));
        const flow = round.sheets.find((s) => s.kind !== "cx")!;
        // The store moved on and no hook reported it.
        flow.data = [["perm", "MISSED"]];
        const healedText = serializeFlow(round);
        await persistReplica(round, healedText);
        const written = parseSidecar(fs.files.get(round.id)!, round.id, hashText(healedText));
        const texts = Object.values(written!.doc.sheets[flow.id].cells).map((c) => c.text);
        expect(texts).toContain("MISSED");
    });

    it("writes nothing while shared editing is off", async () => {
        await recoverReplica(round, serializeFlow(round));
        useFlowStore.setState({ collabEnabled: false });
        await persistReplica(round, serializeFlow(round));
        expect(fs.files.size).toBe(0);
    });

    it("never rejects when the sidecar cannot be written", async () => {
        await recoverReplica(round, serializeFlow(round));
        setSidecarFs({
            read: async () => null,
            write: async () => {
                throw new Error("disk full");
            },
        });
        await expect(persistReplica(round, serializeFlow(round))).resolves.toBeUndefined();
    });

    it("does nothing for a round the replica is not holding", async () => {
        clearReplica();
        await persistReplica(round, serializeFlow(round));
        expect(fs.files.size).toBe(0);
    });

    it("does nothing for a round other than the one open", async () => {
        await recoverReplica(round, serializeFlow(round));
        const other = makeFlowRound({});
        await persistReplica(other, serializeFlow(other));
        expect(fs.files.has(other.id)).toBe(false);
    });
});
