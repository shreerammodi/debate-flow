import { beforeEach, describe, expect, it } from "vitest";

import { seedDoc } from "@/lib/collab/doc";
import { merge } from "@/lib/collab/merge";
import { applyOp, type OpContext } from "@/lib/collab/ops";
import { createMemoryNet } from "@/lib/collab/peerLinkMemory";
import { startCollabSession, type CollabSession } from "@/lib/collab/session";
import { createClock } from "@/lib/collab/stamp";
import { encodeTicket } from "@/lib/collab/ticket";
import type { CollabDoc } from "@/lib/collab/types";
import { makeFlowRound, type FlowRound } from "@/lib/model/flow";
import { useFlowStore } from "@/lib/store/useFlowStore";

const net = createMemoryNet();

/** What iroh hands back. A ticket names the host, so the host holds a real one. */
const HOST = "c".repeat(64);

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
            for (let guard = 0; guard < 50; guard++) {
                const due = pending.filter((p) => p.at <= now);
                if (due.length === 0) return;
                pending = pending.filter((p) => p.at > now);
                for (const p of due) p.fn();
            }
        },
        reset() {
            pending = [];
            now = 0;
        },
    };
}

const clock = manualClock();
let round: FlowRound;
let sheetId: string;

function side(actor: string, startMs: number) {
    let t = startMs;
    const ctx: OpContext = { actor, clock: createClock(actor, () => t++) };
    let doc = seedDoc(round);
    return {
        doc: () => doc,
        apply: (incoming: CollabDoc) => {
            const result = merge(doc, incoming);
            doc = result.doc;
            return result.dropped;
        },
        type: (col: number, row: number, text: string) => {
            doc = applyOp(doc, { kind: "cellText", sheetId, col, row, text }, ctx);
        },
    };
}

/** Two coalesce windows: enough for a push and for the host to pass it on. */
async function settle(): Promise<void> {
    for (let pass = 0; pass < 3; pass++) {
        for (let i = 0; i < 15; i++) await Promise.resolve();
        clock.advance(50);
        for (let i = 0; i < 15; i++) await Promise.resolve();
    }
}

beforeEach(() => {
    net.reset();
    clock.reset();
    useFlowStore.setState({ collabEnabled: true, collabRelayEnabled: true });
    round = makeFlowRound({});
    sheetId = round.sheets.find((s) => s.kind !== "cx")!.id;
    round.sheets.find((s) => s.kind !== "cx")!.data = [];
});

describe("three peers in a star", () => {
    it("passes one guest's typing to the other without waiting for repair", async () => {
        const hostSide = side(HOST, 1_000);
        const host = (await startCollabSession({
            createLink: net.create(HOST),
            roundId: round.id,
            appVersion: "0.11.0",
            doc: hostSide.doc,
            apply: hostSide.apply,
            schedule: clock.schedule,
        }))!;

        const guests: { sess: CollabSession; side: ReturnType<typeof side> }[] = [];
        for (const [i, id] of ["ana", "bo"].entries()) {
            const s = side(id, 5_000 + i * 1_000);
            const sess = (await startCollabSession({
                createLink: net.create(id),
                roundId: round.id,
                appVersion: "0.11.0",
                doc: s.doc,
                apply: s.apply,
                ticket: encodeTicket(await host.share("partner")),
                dial: [HOST],
                schedule: clock.schedule,
            }))!;
            guests.push({ sess, side: s });
            await settle();
        }

        const [ana, bo] = guests;
        ana.side.type(0, 0, "ana typed this");
        ana.sess.notifyLocalChange();
        await settle();

        const reached = (s: ReturnType<typeof side>) =>
            Object.values(s.doc().sheets[sheetId].cells).some((c) => c.text === "ana typed this");

        expect(reached(hostSide), "the host holds it").toBe(true);
        expect(reached(bo.side), "the other guest holds it without a repair tick").toBe(true);

        await host.stop();
        for (const g of guests) await g.sess.stop();
    });

    it("gets there eventually on the repair tick alone", async () => {
        const hostSide = side(HOST, 1_000);
        const host = (await startCollabSession({
            createLink: net.create(HOST),
            roundId: round.id,
            appVersion: "0.11.0",
            doc: hostSide.doc,
            apply: hostSide.apply,
            schedule: clock.schedule,
        }))!;
        const guests: { sess: CollabSession; side: ReturnType<typeof side> }[] = [];
        for (const [i, id] of ["ana", "bo"].entries()) {
            const s = side(id, 5_000 + i * 1_000);
            const sess = (await startCollabSession({
                createLink: net.create(id),
                roundId: round.id,
                appVersion: "0.11.0",
                doc: s.doc,
                apply: s.apply,
                ticket: encodeTicket(await host.share("partner")),
                dial: [HOST],
                schedule: clock.schedule,
            }))!;
            guests.push({ sess, side: s });
            await settle();
        }
        const [ana, bo] = guests;
        ana.side.type(0, 0, "ana typed this");
        ana.sess.notifyLocalChange();
        await settle();

        // Two repair rounds: one to carry it to the host, one onward.
        for (let i = 0; i < 4; i++) {
            clock.advance(5_000);
            for (let j = 0; j < 20; j++) await Promise.resolve();
        }
        expect(
            Object.values(bo.side.doc().sheets[sheetId].cells).some(
                (c) => c.text === "ana typed this",
            ),
        ).toBe(true);

        await host.stop();
        for (const g of guests) await g.sess.stop();
    });
});
