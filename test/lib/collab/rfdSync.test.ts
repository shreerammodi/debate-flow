import { beforeEach, describe, expect, it } from "vitest";

import { seedDoc } from "@/lib/collab/doc";
import { merge } from "@/lib/collab/merge";
import { applyOp, type OpContext } from "@/lib/collab/ops";
import { incomingDoc, LOCAL_RFD_PATH, outgoingDoc, peerNotePath } from "@/lib/collab/rfdSync";
import { createClock } from "@/lib/collab/stamp";
import type { CollabDoc } from "@/lib/collab/types";
import { makeFlowRound, type FlowRound } from "@/lib/model/flow";

const ME = "me-endpoint";
const THEM = "them-endpoint";

let round: FlowRound;
let doc: CollabDoc;
let ctx: OpContext;

beforeEach(() => {
    round = makeFlowRound({});
    doc = seedDoc(round);
    let t = 1_000;
    ctx = { actor: ME, clock: createClock(ME, () => t++) };
});

function withMyNotes(text: string): CollabDoc {
    return applyOp(doc, { kind: "roundField", path: LOCAL_RFD_PATH, value: text }, ctx);
}

describe("outgoingDoc", () => {
    it("sends my notes under my own id, never as the local field", () => {
        const out = outgoingDoc(withMyNotes("voting aff on turns"), ME);
        expect(out.round[peerNotePath(ME)].value).toBe("voting aff on turns");
        expect(out.round[LOCAL_RFD_PATH]).toBeUndefined();
    });

    it("leaves every other field exactly where it was", () => {
        const out = outgoingDoc(withMyNotes("mine"), ME);
        expect(out.round.event).toEqual(doc.round.event);
        expect(out.sheets).toEqual(doc.sheets);
    });

    it("is unchanged for a document with no notes at all", () => {
        expect(outgoingDoc(doc, ME)).toEqual(doc);
    });

    it("can be applied twice without renaming what it already renamed", () => {
        const once = outgoingDoc(withMyNotes("mine"), ME);
        expect(outgoingDoc(once, ME)).toEqual(once);
    });

    it("does not disturb a note that already belongs to a peer", () => {
        const theirs = applyOp(
            doc,
            { kind: "roundField", path: peerNotePath(THEM), value: "theirs" },
            ctx,
        );
        expect(outgoingDoc(theirs, ME).round[peerNotePath(THEM)].value).toBe("theirs");
    });
});

describe("incomingDoc", () => {
    it("drops a peer's local rfd, because their notes are not mine", () => {
        const theirs = applyOp(
            doc,
            { kind: "roundField", path: LOCAL_RFD_PATH, value: "their private notes" },
            { actor: THEM, clock: createClock(THEM, () => 5_000) },
        );
        expect(incomingDoc(theirs).round[LOCAL_RFD_PATH]).toBeUndefined();
    });

    it("keeps their note under their own id", () => {
        const theirs = applyOp(
            doc,
            { kind: "roundField", path: peerNotePath(THEM), value: "voting neg" },
            ctx,
        );
        expect(incomingDoc(theirs).round[peerNotePath(THEM)].value).toBe("voting neg");
    });

    it("is unchanged for a document carrying no notes", () => {
        expect(incomingDoc(doc)).toEqual(doc);
    });
});

describe("the two together", () => {
    it("never lets a partner overwrite my own notes", () => {
        const mine = withMyNotes("mine, written first");
        const theirs = applyOp(
            seedDoc(round),
            { kind: "roundField", path: LOCAL_RFD_PATH, value: "theirs, written later" },
            { actor: THEM, clock: createClock(THEM, () => 9_000) },
        );

        // Their document reaches me the way the sync layer delivers it.
        const merged = merge(mine, incomingDoc(outgoingDoc(theirs, THEM))).doc;

        expect(merged.round[LOCAL_RFD_PATH].value).toBe("mine, written first");
        expect(merged.round[peerNotePath(THEM)].value).toBe("theirs, written later");
    });

    it("gives each side the other's notes and keeps its own", () => {
        const mine = withMyNotes("mine");
        const theirs = applyOp(
            seedDoc(round),
            { kind: "roundField", path: LOCAL_RFD_PATH, value: "theirs" },
            { actor: THEM, clock: createClock(THEM, () => 9_000) },
        );

        const onMyDisk = merge(mine, incomingDoc(outgoingDoc(theirs, THEM))).doc;
        const onTheirDisk = merge(theirs, incomingDoc(outgoingDoc(mine, ME))).doc;

        expect(onMyDisk.round[LOCAL_RFD_PATH].value).toBe("mine");
        expect(onMyDisk.round[peerNotePath(THEM)].value).toBe("theirs");
        expect(onTheirDisk.round[LOCAL_RFD_PATH].value).toBe("theirs");
        expect(onTheirDisk.round[peerNotePath(ME)].value).toBe("mine");
    });
});
