import { describe, expect, it } from "vitest";

import { deltaSince, emptyVector, vectorOf } from "@/lib/collab/delta";
import { projectDoc, projectSheet, seedDoc } from "@/lib/collab/doc";
import { gridPatchFor } from "@/lib/collab/gridPatch";
import { sheetDigest } from "@/lib/collab/hash";
import { merge } from "@/lib/collab/merge";
import { applyOp, type CollabOp, type OpContext } from "@/lib/collab/ops";
import { createClock } from "@/lib/collab/stamp";
import type { CollabDoc } from "@/lib/collab/types";
import { makeFlowRound, makeFlowSheet, type FlowRound } from "@/lib/model/flow";

/**
 * A policy round runs 7 speech columns per flow sheet. A busy elim has a sheet
 * per off-case position plus case, each argued down to a few dozen rows.
 */
const SHEETS = 8;
const COLS = 7;
const ROWS = 60;

function heavyRound(sheets = SHEETS, rows = ROWS): FlowRound {
    const round = makeFlowRound({});
    while (round.sheets.length < sheets) {
        round.sheets.push(
            makeFlowSheet({
                title: `${round.sheets.length}.`,
                group: "neg",
                order: round.sheets.length,
            }),
        );
    }
    for (const sheet of round.sheets) {
        sheet.data = Array.from({ length: rows }, (_, r) =>
            Array.from({ length: COLS }, (_, c) =>
                r % 3 === 0 ? null : `arg ${r}.${c} on ${sheet.title}`,
            ),
        );
        sheet.meta = { "0,0": { bold: true }, "4,2": { card: true } };
    }
    return round;
}

function ctxFor(actor: string, startMs: number): OpContext {
    let t = startMs;
    return { actor, clock: createClock(actor, () => t++) };
}

function cellCount(doc: CollabDoc): number {
    return Object.values(doc.sheets).reduce((n, s) => n + Object.keys(s.cells).length, 0);
}

/**
 * Median of `runs` timings in ms, after a warmup.
 *
 * The warmup is not politeness. Timed cold, the small case in a scaling
 * comparison is mostly JIT, and the ratio against a warm large case reads as
 * superlinear growth that is not there.
 */
function timeMedian(runs: number, fn: () => void): number {
    for (let i = 0; i < 30; i++) fn();
    const seen: number[] = [];
    for (let i = 0; i < runs; i++) {
        const at = performance.now();
        fn();
        seen.push(performance.now() - at);
    }
    seen.sort((a, b) => a - b);
    return seen[Math.floor(seen.length / 2)];
}

/**
 * The fastest of `runs`, after a warmup.
 *
 * For a ratio between two sizes, the floor is the honest statistic: it is the
 * run the scheduler left alone, so it describes the algorithm rather than what
 * else the suite was doing on the other eleven threads. A median is fine for a
 * budget, where the question is what the work usually costs, and wrong here,
 * where two medians taken under different load produce a ratio that is mostly
 * about the load.
 */
function timeFastest(runs: number, fn: () => void): number {
    for (let i = 0; i < 30; i++) fn();
    let best = Infinity;
    for (let i = 0; i < runs; i++) {
        const at = performance.now();
        fn();
        best = Math.min(best, performance.now() - at);
    }
    return best;
}

const report: string[] = [];
function record(label: string, ms: number, budget: number): number {
    report.push(`${label.padEnd(52)} ${ms.toFixed(3).padStart(9)} ms   (budget ${budget})`);
    return ms;
}

describe("collaboration under a full round", () => {
    const round = heavyRound();
    const doc = seedDoc(round);
    const sheetId = round.sheets[1].id;
    const alex = ctxFor("alex", 1_000_000);

    it("holds a round of the size it claims to", () => {
        expect(cellCount(doc)).toBe(SHEETS * COLS * ROWS);
        report.push(`\ncells: ${cellCount(doc)}  sheets: ${SHEETS}  cols: ${COLS}  rows: ${ROWS}`);
    });

    // --- The keystroke path --------------------------------------------------

    it("writes a cell that already exists in constant-ish time", () => {
        const ms = timeMedian(50, () => {
            applyOp(doc, { kind: "cellText", sheetId, col: 3, row: 30, text: "retyped" }, alex);
        });
        expect(record("applyOp cellText, existing cell", ms, 2)).toBeLessThan(2);
    });

    it("grows a column to a fresh row without scanning the sheet per blank", () => {
        const empty = seedDoc(makeFlowRound({}));
        const freshId = Object.keys(empty.sheets)[1];
        const ms = timeMedian(20, () => {
            applyOp(
                empty,
                { kind: "cellText", sheetId: freshId, col: 0, row: 80, text: "deep" },
                alex,
            );
        });
        expect(record("applyOp cellText, grows 80 blanks", ms, 25)).toBeLessThan(25);
    });

    it("inserts a row across every column", () => {
        const ms = timeMedian(20, () => {
            applyOp(doc, { kind: "insertRow", sheetId, row: 20 }, alex);
        });
        expect(record("applyOp insertRow across 7 columns", ms, 20)).toBeLessThan(20);
    });

    // --- The push path, once per typing burst --------------------------------

    it("computes a delta the far side has already seen", () => {
        const seen = vectorOf(doc);
        const ms = timeMedian(30, () => deltaSince(doc, seen));
        expect(record("deltaSince, nothing to send", ms, 20)).toBeLessThan(20);
    });

    it("computes a delta carrying one new cell", () => {
        const seen = vectorOf(doc);
        const typed = applyOp(
            doc,
            { kind: "cellText", sheetId, col: 0, row: 0, text: "new" },
            alex,
        );
        const ms = timeMedian(30, () => {
            const delta = deltaSince(typed, seen);
            expect(Object.keys(delta.sheets)).toHaveLength(1);
        });
        expect(record("deltaSince, one new cell", ms, 20)).toBeLessThan(20);
    });

    it("suppresses the whole seed against a fresh peer's vector", () => {
        const ms = timeMedian(30, () => {
            const delta = deltaSince(doc, emptyVector());
            expect(delta.sheets).toEqual({});
        });
        expect(record("deltaSince, origin vector (first sync)", ms, 20)).toBeLessThan(20);
    });

    it("builds the repair vector", () => {
        const ms = timeMedian(30, () => vectorOf(doc));
        expect(record("vectorOf, repair tick", ms, 20)).toBeLessThan(20);
    });

    // --- The inbound path, once per remote change ----------------------------

    it("merges a one-cell delta into a full round", () => {
        const sam = ctxFor("sam", 5_000_000);
        const delta = deltaSince(
            applyOp(doc, { kind: "cellText", sheetId, col: 1, row: 5, text: "theirs" }, sam),
            vectorOf(doc),
        );
        const ms = timeMedian(30, () => merge(doc, delta));
        expect(record("merge, one-cell delta", ms, 20)).toBeLessThan(20);
    });

    it("merges a whole document, which is what a reconnect sends", () => {
        const ms = timeMedian(10, () => merge(doc, doc));
        expect(record("merge, full state (reconnect)", ms, 60)).toBeLessThan(60);
    });

    it("projects the round the store is handed", () => {
        const ms = timeMedian(20, () => projectDoc(doc, round));
        expect(record("projectDoc, cold", ms, 60)).toBeLessThan(60);
    });

    // What actually runs per remote change. A delta names only the sheets it
    // has something new for, so a merge builds one new sheet object and the
    // other seven are handed straight back.
    it("projects only the sheet a partner touched", () => {
        const sam = ctxFor("sam", 7_000_000);
        const base = projectDoc(doc, round);
        const delta = deltaSince(
            applyOp(doc, { kind: "cellText", sheetId, col: 0, row: 1, text: "theirs" }, sam),
            vectorOf(doc),
        );
        expect(Object.keys(delta.sheets)).toHaveLength(1);
        const after = merge(doc, delta).doc;

        const whole = timeMedian(20, () => projectDoc(after, base));
        const reused = timeMedian(20, () => projectDoc(after, base, doc));
        report.push(
            `projectDoc, 1 of ${SHEETS} sheets changed`.padEnd(52) +
                `${reused.toFixed(3).padStart(9)} ms   (whole round ${whole.toFixed(3)})`,
        );
        // Seven sheets of eight are handed back, so most of the work is gone.
        expect(reused).toBeLessThan(whole * 0.6);
    });

    it("diffs one sheet into the writes a pane needs", () => {
        const sam = ctxFor("sam", 6_000_000);
        const after = applyOp(
            doc,
            { kind: "cellText", sheetId, col: 2, row: 9, text: "theirs" },
            sam,
        );
        const ms = timeMedian(30, () => {
            const patch = gridPatchFor(doc.sheets[sheetId], after.sheets[sheetId]);
            expect(patch.writes).toHaveLength(1);
        });
        expect(record("gridPatchFor, one changed cell", ms, 20)).toBeLessThan(20);
    });

    // --- The save path, once per autosave ------------------------------------

    it("digests every sheet the way the drift check does", () => {
        const ms = timeMedian(20, () => {
            for (const sheet of Object.values(doc.sheets)) {
                const p = projectSheet(sheet);
                sheetDigest(p.data, p.meta);
            }
        });
        expect(record("sheetDigest x8, every autosave", ms, 80)).toBeLessThan(80);
    });

    it("prints the numbers", () => {
        console.log(`\n${report.join("\n")}\n`);
    });
});

describe("cost as a round grows", () => {
    /**
     * Doubling the cells must not quadruple the work on any hot path.
     *
     * Both points are large on purpose. A small round projects in tens of
     * microseconds, where scheduling jitter is a large fraction of the
     * measurement and the ratio against it says more about the timer than
     * about the algorithm.
     */
    it("stays linear in the number of cells", () => {
        const scale = (rows: number) => {
            const r = heavyRound(4, rows);
            const d = seedDoc(r);
            const seen = vectorOf(d);
            const alex = ctxFor("alex", 2_000_000);
            const sheetId = r.sheets[1].id;
            return {
                cells: cellCount(d),
                delta: timeFastest(30, () => deltaSince(d, seen)),
                merge: timeFastest(30, () => merge(d, d)),
                project: timeFastest(30, () => projectDoc(d, r)),
                write: timeFastest(50, () =>
                    applyOp(d, { kind: "cellText", sheetId, col: 3, row: 10, text: "x" }, alex),
                ),
            };
        };

        const small = scale(120);
        const large = scale(240);
        const growth = large.cells / small.cells;
        const lines = [`\ncells ${small.cells} -> ${large.cells} (x${growth})`];
        const factors: Record<string, number> = {};
        for (const key of ["delta", "merge", "project", "write"] as const) {
            factors[key] = large[key] / Math.max(small[key], 0.001);
            lines.push(
                `${key.padEnd(10)} ${small[key].toFixed(3)} -> ${large[key].toFixed(3)} ms  x${factors[key].toFixed(2)}`,
            );
        }
        console.log(`${lines.join("\n")}\n`);
        // Linear is x2. Quadratic would be x4, and that is the thing to catch.
        for (const [key, factor] of Object.entries(factors)) {
            expect(factor, `${key} grew x${factor.toFixed(2)} for x${growth} cells`).toBeLessThan(
                3.5,
            );
        }
    });
});
