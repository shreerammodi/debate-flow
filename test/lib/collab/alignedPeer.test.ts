import { beforeEach, describe, expect, it } from "vitest";

import { projectDoc, seedDoc } from "@/lib/collab/doc";
import { merge } from "@/lib/collab/merge";
import { applyOp, type OpContext } from "@/lib/collab/ops";
import { createMemoryNet } from "@/lib/collab/peerLinkMemory";
import { startCollabSession, type CollabSession } from "@/lib/collab/session";
import { createClock } from "@/lib/collab/stamp";
import { encodeTicket } from "@/lib/collab/ticket";
import type { CollabDoc } from "@/lib/collab/types";
import {
    gridCol,
    modelCol,
    toGridCol,
    toModelCol,
    type GridCol,
    type ModelCol,
} from "@/lib/grid/colSpace";
import { columnsForFlowSheet, spacerColumns, spacerCount } from "@/lib/grid/flowColumns";
import { getPresences } from "@/lib/grid/presenceBridge";
import { makeFlowRound, makeFlowSheet, type FlowRound, type FlowSheet } from "@/lib/model/flow";
import { useFlowStore } from "@/lib/store/useFlowStore";

const net = createMemoryNet();

/** What iroh hands back. A ticket names the host, so the host holds a real one. */
const ALIGNED = "a".repeat(64);
const PLAIN = "b".repeat(64);

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
let sheet: FlowSheet;

/** A replica the session can read and write, with no grid behind it. */
function replica(actor: string, startMs: number) {
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
        type: (col: ModelCol, row: number, text: string) => {
            doc = applyOp(doc, { kind: "cellText", sheetId: sheet.id, col, row, text }, ctx);
        },
        /** The sheet as this side's file would hold it. */
        sheet: () => projectDoc(doc, round).sheets.find((s) => s.id === sheet.id)!,
    };
}

/**
 * One debater's view of the shared sheet: the columns their pane draws and the
 * conversion it runs at every seam.
 *
 * Alignment is a display setting and is not synced, so it is a parameter here
 * rather than a store read: one process holds one store, and the whole point
 * of this file is two peers holding different answers at once.
 */
function pane(aligned: boolean) {
    const spacers = spacerCount(round, sheet.id, aligned);
    const shown = columnsForFlowSheet(round, sheet);
    // What the pane draws: the pad, then the sheet's own columns.
    const cols = aligned ? [...spacerColumns(round, sheet), ...shown] : shown;
    return {
        spacers,
        /** Where this pane draws a speech, which is where its debater clicks. */
        clicks: (speechId: string): GridCol => gridCol(cols.findIndex((c) => c.id === speechId)),
        /**
         * The speech printed above a grid column. What the debater is looking
         * at, taken off the drawn header row and through no conversion, so it
         * stands as the intent a conversion is held against.
         */
        drawnAt: (col: GridCol): string | undefined => cols[col]?.id,
        /** The cell a grid column of this pane points at, or null in the pad. */
        cellAt: (col: GridCol): ModelCol | null => toModelCol(col, spacers),
        /** The speech this pane shows a model column under. */
        speechOf: (col: ModelCol): string | undefined => cols[toGridCol(col, spacers)]?.id,
    };
}

/** Two coalesce windows: enough for a push and for the far side to take it. */
async function settle(): Promise<void> {
    for (let pass = 0; pass < 3; pass++) {
        for (let i = 0; i < 15; i++) await Promise.resolve();
        clock.advance(50);
        for (let i = 0; i < 15; i++) await Promise.resolve();
    }
}

/** The two connected sides, aligned as host and unaligned as its guest. */
async function link() {
    const alignedSide = replica(ALIGNED, 1_000);
    const host = (await startCollabSession({
        createLink: net.create(ALIGNED),
        roundId: round.id,
        appVersion: "0.11.0",
        doc: alignedSide.doc,
        apply: alignedSide.apply,
        schedule: clock.schedule,
    }))!;
    const plainSide = replica(PLAIN, 5_000);
    const guest = (await startCollabSession({
        createLink: net.create(PLAIN),
        roundId: round.id,
        appVersion: "0.11.0",
        doc: plainSide.doc,
        apply: plainSide.apply,
        ticket: encodeTicket(await host.share("partner")),
        dial: [ALIGNED],
        schedule: clock.schedule,
    }))!;
    await settle();
    return { host, guest, alignedSide, plainSide };
}

async function hangUp(...sessions: CollabSession[]): Promise<void> {
    for (const s of sessions) await s.stop();
}

/** The model column a row of a projected sheet holds `text` in. */
function landedAt(data: (string | null)[][], row: number, text: string): ModelCol {
    const col = data[row]?.indexOf(text) ?? -1;
    expect(col, "the partner holds the cell at all").toBeGreaterThanOrEqual(0);
    return modelCol(col);
}

beforeEach(() => {
    net.reset();
    clock.reset();
    useFlowStore.setState({ collabEnabled: true, collabRelayEnabled: true });
    round = makeFlowRound({});
    // A neg sheet opening on the block. Three speeches of the Policy order sit
    // left of it, so an aligned pane leads with three spacers where an
    // unaligned one leads with none: every column of this sheet is at a
    // different screen position on the two peers.
    sheet = { ...makeFlowSheet({ title: "2.", group: "neg", order: 1 }), startSpeechId: "block" };
    round.sheets.push(sheet);
});

/**
 * Two debaters on one round whose alignment settings differ. Nobody tests this
 * by hand: it takes two machines and two people. A column index on the wire is
 * a model column, and a pane that converted it wrong would put a partner's
 * argument under a different speech, silently, in both files.
 *
 * This drives the conversion at the seam rather than through a mounted pane,
 * because two panes cannot hold two different alignment settings out of one
 * store. That HotGrid runs these conversions on every one of its own seams is
 * what `test/components/flow/HotGridColSpace.test.tsx` sweeps.
 */
describe("a padded pane and an unpadded peer", () => {
    it("draws one speech at two screen columns at once", () => {
        const aligned = pane(true);
        const plain = pane(false);
        expect(aligned.spacers).toBe(3);
        expect(plain.spacers).toBe(0);
        // The premise: every column of this sheet is somewhere else on the
        // two peers, so a conversion that dropped out would be visible.
        expect(aligned.clicks("1ar")).not.toBe(plain.clicks("1ar"));
        expect(aligned.drawnAt(aligned.clicks("1ar"))).toBe("1ar");
        expect(plain.drawnAt(plain.clicks("1ar"))).toBe("1ar");
    });

    it("lands a cell the aligned side typed on the same speech for its peer", async () => {
        const { host, guest, alignedSide, plainSide } = await link();
        const aligned = pane(true);
        const plain = pane(false);

        const clicked = aligned.clicks("1ar");
        expect(aligned.drawnAt(clicked), "the debater is under the 1AR header").toBe("1ar");
        const at = aligned.cellAt(clicked);
        expect(at, "the 1AR column is a cell, not a spacer").not.toBeNull();
        alignedSide.type(at!, 0, "perm do both");
        host.notifyLocalChange();
        await settle();

        const landed = landedAt(plainSide.sheet().data, 0, "perm do both");
        expect(plain.speechOf(landed)).toBe(aligned.drawnAt(clicked));

        await hangUp(host, guest);
    });

    it("lands a cell the unaligned side typed on the same speech for its peer", async () => {
        const { host, guest, alignedSide, plainSide } = await link();
        const aligned = pane(true);
        const plain = pane(false);

        const clicked = plain.clicks("2nr");
        expect(plain.drawnAt(clicked), "the debater is under the 2NR header").toBe("2nr");
        const at = plain.cellAt(clicked);
        expect(at, "the 2NR column is a cell, not a spacer").not.toBeNull();
        plainSide.type(at!, 2, "extend the kritik");
        guest.notifyLocalChange();
        await settle();

        const landed = landedAt(alignedSide.sheet().data, 2, "extend the kritik");
        expect(aligned.speechOf(landed)).toBe(plain.drawnAt(clicked));

        await hangUp(host, guest);
    });

    it("names the speech the aligned side's cursor is really in", async () => {
        const { host, guest } = await link();
        const aligned = pane(true);
        const plain = pane(false);

        const clicked = aligned.clicks("block");
        expect(aligned.drawnAt(clicked), "the debater is under the Block header").toBe("block");
        const at = aligned.cellAt(clicked);
        expect(at, "the block column is a cell, not a spacer").not.toBeNull();
        host.setCursor({ sheetId: sheet.id, col: at!, row: 4 });
        await settle();

        // The unaligned side's session is the only one that took a position
        // message, so the table is its reading of where its partner is.
        const seen = getPresences();
        expect(seen).toHaveLength(1);
        expect(seen[0].endpointId).toBe(ALIGNED);
        expect(seen[0].sheetId).toBe(sheet.id);
        expect(seen[0].row).toBe(4);
        expect(plain.speechOf(seen[0].col)).toBe(aligned.drawnAt(clicked));

        await hangUp(host, guest);
    });
});
