import { beforeEach, describe, expect, it } from "vitest";

import { seedDoc } from "@/lib/collab/doc";
import { merge } from "@/lib/collab/merge";
import { applyOp, type OpContext } from "@/lib/collab/ops";
import {
    dropSelfNote,
    incomingDoc,
    LOCAL_RFD_PATH,
    outgoingDoc,
    peerNotePath,
} from "@/lib/collab/rfdSync";
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
        expect(incomingDoc(theirs, ME, THEM).round[LOCAL_RFD_PATH]).toBeUndefined();
    });

    it("keeps their note under their own id", () => {
        const theirs = applyOp(
            doc,
            { kind: "roundField", path: peerNotePath(THEM), value: "voting neg" },
            ctx,
        );
        expect(incomingDoc(theirs, ME, THEM).round[peerNotePath(THEM)].value).toBe("voting neg");
    });

    // Every peer holds my note under my id, because that is how I sent it. It
    // must not come home: beside my own rfd it reads as a partner's reasoning,
    // and nothing on this machine ever updates it again.
    it("drops my own note echoed back under my own id", () => {
        const echoed = applyOp(
            doc,
            { kind: "roundField", path: peerNotePath(ME), value: "what they last saw me write" },
            { actor: THEM, clock: createClock(THEM, () => 5_000) },
        );
        expect(incomingDoc(echoed, ME, THEM).round[peerNotePath(ME)]).toBeUndefined();
    });

    // Nothing signs a register, so a note is the sender's or it is nobody's.
    // Otherwise a peer writes under a partner's or the judge's id and the
    // drawer, the print view and the export all label it with that name.
    it("refuses a note the sender filed under a third party's id", () => {
        const JUDGE = "judge-endpoint";
        const forged = applyOp(
            doc,
            { kind: "roundField", path: peerNotePath(JUDGE), value: "voting neg, obviously" },
            { actor: THEM, clock: createClock(THEM, () => 5_000) },
        );
        const applied = incomingDoc(forged, ME, THEM);
        expect(applied.round[peerNotePath(JUDGE)]).toBeUndefined();
        // And the sender cannot smuggle it through by sending both at once.
        const both = applyOp(
            forged,
            { kind: "roundField", path: peerNotePath(THEM), value: "mine, honestly" },
            { actor: THEM, clock: createClock(THEM, () => 5_001) },
        );
        const merged = incomingDoc(both, ME, THEM);
        expect(merged.round[peerNotePath(JUDGE)]).toBeUndefined();
        expect(merged.round[peerNotePath(THEM)].value).toBe("mine, honestly");
    });

    it("is unchanged for a document carrying no notes", () => {
        expect(incomingDoc(doc, ME, THEM)).toEqual(doc);
    });
});

describe("dropSelfNote", () => {
    it("takes my own note back out of my own document", () => {
        const poisoned = applyOp(
            doc,
            { kind: "roundField", path: peerNotePath(ME), value: "stale copy of mine" },
            ctx,
        );
        expect(dropSelfNote(poisoned, ME).round[peerNotePath(ME)]).toBeUndefined();
    });

    it("leaves my own rfd and a real peer's note alone", () => {
        const both = applyOp(
            withMyNotes("mine"),
            { kind: "roundField", path: peerNotePath(THEM), value: "theirs" },
            ctx,
        );
        const clean = dropSelfNote(both, ME);
        expect(clean.round[LOCAL_RFD_PATH].value).toBe("mine");
        expect(clean.round[peerNotePath(THEM)].value).toBe("theirs");
    });

    it("hands back the same document when there is nothing to drop", () => {
        expect(dropSelfNote(doc, ME)).toBe(doc);
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
        const merged = merge(mine, incomingDoc(outgoingDoc(theirs, THEM), ME, THEM)).doc;

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

        const onMyDisk = merge(mine, incomingDoc(outgoingDoc(theirs, THEM), ME, THEM)).doc;
        const onTheirDisk = merge(theirs, incomingDoc(outgoingDoc(mine, ME), THEM, ME)).doc;

        expect(onMyDisk.round[LOCAL_RFD_PATH].value).toBe("mine");
        expect(onMyDisk.round[peerNotePath(THEM)].value).toBe("theirs");
        expect(onTheirDisk.round[LOCAL_RFD_PATH].value).toBe("theirs");
        expect(onTheirDisk.round[peerNotePath(ME)].value).toBe("mine");
    });

    // The round trip that put a debater's own words on screen as a partner's:
    // I send, they hold it under my id, they send their whole state back.
    it("leaves nothing of mine behind when my own document comes home", () => {
        const mine = withMyNotes("123123");
        const theirs = merge(seedDoc(round), outgoingDoc(mine, ME)).doc;
        expect(theirs.round[peerNotePath(ME)].value).toBe("123123");

        const back = merge(mine, incomingDoc(outgoingDoc(theirs, THEM), ME, THEM)).doc;
        expect(back.round[peerNotePath(ME)]).toBeUndefined();
        expect(back.round[LOCAL_RFD_PATH].value).toBe("123123");
    });
});
