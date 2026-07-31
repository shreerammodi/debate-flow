import { beforeEach, describe, expect, it } from "vitest";

import { deltaSince, vectorOf } from "@/lib/collab/delta";
import { seedDoc } from "@/lib/collab/doc";
import { merge } from "@/lib/collab/merge";
import { applyOp, type OpContext } from "@/lib/collab/ops";
import type { PeerConn, WireMessage } from "@/lib/collab/peerLink";
import { createClock } from "@/lib/collab/stamp";
import { attachSync } from "@/lib/collab/sync";
import type { CollabDoc } from "@/lib/collab/types";
import { makeFlowRound, type FlowRound } from "@/lib/model/flow";

/** A connection that records what was sent and lets a test deliver inbound. */
function fakeConn() {
    const sent: WireMessage[] = [];
    let onMsg: ((m: WireMessage) => void) | null = null;
    let onClose: (() => void) | null = null;
    const conn: PeerConn = {
        id: "sam",
        connectionType: () => "direct",
        relayUrl: () => null,
        send: (m) => sent.push(m),
        onMessage: (cb) => {
            onMsg = cb;
        },
        onClose: (cb) => {
            onClose = cb;
        },
        close: () => onClose?.(),
    };
    return {
        conn,
        sent,
        deliver: (m: WireMessage) => onMsg?.(m),
    };
}

/** A scheduler the test steps by hand, so no timer library is needed. */
function manualClock() {
    let pending: { fn: () => void; at: number }[] = [];
    let now = 0;
    return {
        schedule(fn: () => void, ms: number) {
            const entry = { fn, at: now + ms };
            pending.push(entry);
            return () => {
                pending = pending.filter((p) => p !== entry);
            };
        },
        advance(ms: number) {
            now += ms;
            const due = pending.filter((p) => p.at <= now);
            pending = pending.filter((p) => p.at > now);
            for (const p of due) p.fn();
        },
        get pendingCount() {
            return pending.length;
        },
    };
}

let round: FlowRound;
let sheetId: string;
let doc: CollabDoc;
let alex: OpContext;

beforeEach(() => {
    round = makeFlowRound({});
    const flow = round.sheets.find((s) => s.kind !== "cx")!;
    sheetId = flow.id;
    flow.data = [
        ["perm", "link"],
        ["cap bad", "turn"],
    ];
    doc = seedDoc(round);
    let t = 1_000;
    alex = { actor: "alex", clock: createClock("alex", () => t++) };
});

function edit(text: string, col = 0, row = 0): void {
    doc = applyOp(doc, { kind: "cellText", sheetId, col, row, text }, alex);
}

function setup(over: { readOnly?: boolean } = {}) {
    const link = fakeConn();
    const clock = manualClock();
    const applied: CollabDoc[] = [];
    const sync = attachSync({
        conn: link.conn,
        doc: () => doc,
        apply: (incoming) => {
            applied.push(incoming);
            doc = merge(doc, incoming).doc;
            return [];
        },
        readOnly: over.readOnly,
        endpointId: "me",
        from: "them",
        schedule: clock.schedule,
    });
    return { ...link, clock, applied, sync };
}

describe("pushing local changes", () => {
    it("coalesces a burst into one delta", () => {
        const s = setup();
        edit("one");
        s.sync.notifyLocalChange();
        edit("two");
        s.sync.notifyLocalChange();
        edit("three");
        s.sync.notifyLocalChange();
        expect(s.sent).toHaveLength(0);

        s.clock.advance(30);
        const deltas = s.sent.filter((m) => m.type === "delta");
        expect(deltas).toHaveLength(1);
    });

    it("sends only the cell that changed", () => {
        const s = setup();
        edit("changed", 1, 1);
        s.sync.notifyLocalChange();
        s.clock.advance(30);
        const delta = s.sent.find((m) => m.type === "delta")!;
        if (delta.type !== "delta") throw new Error("expected a delta");
        const cells = Object.values(delta.doc.sheets[sheetId].cells);
        expect(cells).toHaveLength(1);
        expect(cells[0].text).toBe("changed");
    });

    it("sends nothing when there is nothing new", () => {
        const s = setup();
        s.sync.notifyLocalChange();
        s.clock.advance(30);
        expect(s.sent.filter((m) => m.type === "delta")).toHaveLength(0);
    });

    it("does not resend what it already sent", () => {
        const s = setup();
        edit("one");
        s.sync.notifyLocalChange();
        s.clock.advance(30);
        s.sync.notifyLocalChange();
        s.clock.advance(30);
        expect(s.sent.filter((m) => m.type === "delta")).toHaveLength(1);
    });
});

describe("receiving", () => {
    it("applies an inbound delta", () => {
        const s = setup();
        const theirs = applyOp(
            seedDoc(round),
            { kind: "cellText", sheetId, col: 1, row: 0, text: "from sam" },
            { actor: "sam", clock: createClock("sam", () => 9_000) },
        );
        s.deliver({ type: "delta", doc: deltaSince(theirs, vectorOf(doc)) });
        expect(s.applied).toHaveLength(1);
        const cell = Object.values(doc.sheets[sheetId].cells).find((c) => c.text === "from sam");
        expect(cell).toBeDefined();
    });

    it("answers a vector with exactly what the far side is missing", () => {
        const s = setup();
        const theirVector = vectorOf(doc);
        edit("after their vector");
        s.deliver({ type: "vector", seen: theirVector });
        const reply = s.sent.find((m) => m.type === "delta");
        if (!reply || reply.type !== "delta") throw new Error("expected a delta reply");
        const cells = Object.values(reply.doc.sheets[sheetId].cells);
        expect(cells).toHaveLength(1);
        expect(cells[0].text).toBe("after their vector");
    });

    it("answers a vector that is already current with nothing", () => {
        const s = setup();
        s.deliver({ type: "vector", seen: vectorOf(doc) });
        expect(s.sent.filter((m) => m.type === "delta")).toHaveLength(0);
    });

    it("merges an inbound state and replies with what the sender lacks", () => {
        const s = setup();
        edit("only here");
        const theirs = seedDoc(round);
        s.deliver({ type: "state", doc: theirs });
        expect(s.applied).toHaveLength(1);
        const reply = s.sent.find((m) => m.type === "delta");
        expect(reply).toBeDefined();
    });
});

describe("a read-only peer", () => {
    it("never reaches apply, because the host enforces the role", () => {
        const s = setup({ readOnly: true });
        const theirs = applyOp(
            seedDoc(round),
            { kind: "cellText", sheetId, col: 0, row: 0, text: "coach typed" },
            { actor: "coach", clock: createClock("coach", () => 9_000) },
        );
        s.deliver({ type: "delta", doc: theirs });
        s.deliver({ type: "state", doc: theirs });
        expect(s.applied).toHaveLength(0);
    });

    it("still receives what the host sends", () => {
        const s = setup({ readOnly: true });
        edit("host typed");
        s.sync.notifyLocalChange();
        s.clock.advance(30);
        expect(s.sent.filter((m) => m.type === "delta")).toHaveLength(1);
    });
});

describe("repair", () => {
    it("sends a vector on its own schedule, as repair and not as the path", () => {
        const s = setup();
        s.clock.advance(5_000);
        expect(s.sent.filter((m) => m.type === "vector").length).toBeGreaterThan(0);
    });
});

describe("stop", () => {
    it("cancels a pending push so a torn-down session cannot write", () => {
        const s = setup();
        edit("one");
        s.sync.notifyLocalChange();
        s.sync.stop();
        s.clock.advance(1_000);
        expect(s.sent.filter((m) => m.type === "delta")).toHaveLength(0);
    });

    it("leaves no timer behind", () => {
        const s = setup();
        s.sync.notifyLocalChange();
        s.sync.stop();
        expect(s.clock.pendingCount).toBe(0);
    });
});

describe("sendState", () => {
    it("sends the whole document for a peer joining with no file", () => {
        const s = setup();
        s.sync.sendState();
        const state = s.sent.find((m) => m.type === "state");
        if (!state || state.type !== "state") throw new Error("expected a state");
        expect(Object.keys(state.doc.sheets[sheetId].cells)).toHaveLength(4);
    });

    it("does not then resend the same document as a delta", () => {
        const s = setup();
        s.sync.sendState();
        s.sync.notifyLocalChange();
        s.clock.advance(30);
        expect(s.sent.filter((m) => m.type === "delta")).toHaveLength(0);
    });

    it("still ships an edit made after the state", () => {
        const s = setup();
        s.sync.sendState();
        edit("after state");
        s.sync.notifyLocalChange();
        s.clock.advance(30);
        expect(s.sent.filter((m) => m.type === "delta")).toHaveLength(1);
    });
});
