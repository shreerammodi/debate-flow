/**
 * What an admitted peer can do with the bytes it is allowed to send.
 *
 * Every case here goes in through `parseWireMessage`, which is the only runtime
 * validator on the peer path, so a test that passes proves the shipping
 * transport would behave the same way. The adversary is a peer that has been
 * admitted once and runs a modified client: it chooses every byte below.
 */

import { describe, expect, it } from "vitest";

import { vectorOf } from "@/lib/collab/delta";
import { projectDoc, seedDoc, sheetWidth } from "@/lib/collab/doc";
import { merge } from "@/lib/collab/merge";
import { applyOp, type OpContext } from "@/lib/collab/ops";
import { parseWireMessage, type WireMessage } from "@/lib/collab/peerLink";
import { createClock, type Stamp } from "@/lib/collab/stamp";
import { cellKey, type CollabCell, type CollabDoc } from "@/lib/collab/types";
import { columnsForFlowSheet } from "@/lib/grid/flowColumns";
import { makeFlowRound, type FlowRound } from "@/lib/model/flow";

/** Far above any stamp the victim's own clock has reached, and a safe integer. */
const HIGH: Stamp = { ms: 9_000_000_000_000, counter: 0, actor: "attacker" };

function roundWithData(): FlowRound {
    const round = makeFlowRound({});
    const flow = round.sheets.find((s) => s.kind !== "cx")!;
    flow.data = [
        ["perm do both", "no link"],
        ["cap bad", null],
    ];
    return round;
}

/** A line exactly as it arrives: JSON text, then the transport's own parser. */
function onWire(json: string): WireMessage | null {
    return parseWireMessage(JSON.parse(json));
}

/** The document out of a `delta` the transport accepted. Throws if it refused. */
function inboundDoc(json: string): CollabDoc {
    const msg = onWire(json);
    if (!msg || (msg.type !== "delta" && msg.type !== "state")) {
        throw new Error("the transport refused this line");
    }
    return msg.doc;
}

function deltaOf(doc: unknown): CollabDoc {
    return inboundDoc(JSON.stringify({ type: "delta", doc }));
}

function hostileCell(over: Partial<CollabCell>): CollabCell {
    return {
        col: 0,
        rank: "1",
        actor: "attacker",
        text: "theirs",
        textStamp: HIGH,
        meta: {},
        metaStamp: HIGH,
        deleted: null,
        ...over,
    };
}

describe("a register value that indexes a static table", () => {
    it("does not make the round fatal to render", () => {
        const round = roundWithData();
        const flow = round.sheets.find((s) => s.kind !== "cx")!;
        // "constructor" is the worst of the family: it reaches Object.prototype,
        // so an `in` check passes it and the file parser lets it back in too.
        const incoming = deltaOf({
            roundId: round.id,
            round: { event: { value: "constructor", stamp: HIGH } },
            sheets: {
                [flow.id]: {
                    id: flow.id,
                    deleted: null,
                    cells: {},
                    fields: { group: { value: "zzz", stamp: HIGH } },
                },
            },
        });

        const poisoned = projectDoc(merge(seedDoc(round), incoming).doc, round);
        const sheet = poisoned.sheets.find((s) => s.id === flow.id)!;
        expect(() => columnsForFlowSheet(poisoned, sheet)).not.toThrow();
        expect(columnsForFlowSheet(poisoned, sheet).length).toBeGreaterThan(0);
    });

    it("does not make the cross-ex sheet fatal either", () => {
        const round = roundWithData();
        const cx = round.sheets.find((s) => s.kind === "cx")!;
        const incoming = deltaOf({
            roundId: round.id,
            round: { firstSide: { value: "nobody", stamp: HIGH } },
            sheets: {},
        });

        const poisoned = projectDoc(merge(seedDoc(round), incoming).doc, round);
        const sheet = poisoned.sheets.find((s) => s.id === cx.id)!;
        expect(() => columnsForFlowSheet(poisoned, sheet)).not.toThrow();
    });
});

describe("a cell's column index", () => {
    it("never becomes the bound of a loop nobody can afford", () => {
        const round = roundWithData();
        const flow = round.sheets.find((s) => s.kind !== "cx")!;
        const incoming = deltaOf({
            roundId: round.id,
            round: {},
            sheets: {
                [flow.id]: {
                    id: flow.id,
                    fields: {},
                    deleted: null,
                    cells: { far: hostileCell({ col: 1e15, rank: "21" }) },
                },
            },
        });

        const merged = merge(seedDoc(round), incoming).doc;
        // The cell is held, so the two replicas still agree; it just projects
        // nowhere, which is what a cell nobody can see already looks like.
        expect(merged.sheets[flow.id].cells.far).toBeDefined();
        expect(sheetWidth(merged.sheets[flow.id])).toBe(2);

        const projected = projectDoc(merged, round);
        const sheet = projected.sheets.find((s) => s.id === flow.id)!;
        expect(sheet.data.every((row) => row.length === 2)).toBe(true);
    });

    it("is ignored when it is not an index at all", () => {
        const round = roundWithData();
        const flow = round.sheets.find((s) => s.kind !== "cx")!;
        const incoming = deltaOf({
            roundId: round.id,
            round: {},
            sheets: {
                [flow.id]: {
                    id: flow.id,
                    fields: {},
                    deleted: null,
                    cells: {
                        fraction: hostileCell({ col: 2.5, rank: "21" }),
                        negative: hostileCell({ col: -1, rank: "31" }),
                    },
                },
            },
        });
        expect(sheetWidth(merge(seedDoc(round), incoming).doc.sheets[flow.id])).toBe(2);
    });
});

describe("a rank this build cannot order", () => {
    function typeBelowLastRow(doc: CollabDoc, sheetId: string): CollabDoc {
        const ctx: OpContext = { actor: "me", clock: createClock("me", () => 5_000) };
        return applyOp(doc, { kind: "cellText", sheetId, col: 0, row: 6, text: "mine" }, ctx);
    }

    it("never joins the round, so the column stays writable", () => {
        const round = roundWithData();
        const flow = round.sheets.find((s) => s.kind !== "cx")!;
        // A trailing zero digit is the one shape `rankBetween` refuses to
        // subdivide, and it sorts to the bottom of the column.
        const incoming = deltaOf({
            roundId: round.id,
            round: {},
            sheets: {
                [flow.id]: {
                    id: flow.id,
                    fields: {},
                    deleted: null,
                    cells: { bottom: hostileCell({ rank: "zzzzz0" }) },
                },
            },
        });

        const merged = merge(seedDoc(round), incoming).doc;
        expect(merged.sheets[flow.id].cells.bottom).toBeUndefined();
        expect(() => typeBelowLastRow(merged, flow.id)).not.toThrow();
    });

    it("is refused when it is long enough to outrun the stack", () => {
        const round = roundWithData();
        const flow = round.sheets.find((s) => s.kind !== "cx")!;
        const incoming = deltaOf({
            roundId: round.id,
            round: {},
            sheets: {
                [flow.id]: {
                    id: flow.id,
                    fields: {},
                    deleted: null,
                    cells: {
                        long: hostileCell({ rank: `${"z".repeat(200_000)}1` }),
                        empty: hostileCell({ rank: "" }),
                        notADigit: hostileCell({ rank: "../../etc/passwd" }),
                    },
                },
            },
        });

        const merged = merge(seedDoc(round), incoming).doc;
        expect(Object.keys(merged.sheets[flow.id].cells)).not.toContain("long");
        expect(Object.keys(merged.sheets[flow.id].cells)).not.toContain("empty");
        expect(Object.keys(merged.sheets[flow.id].cells)).not.toContain("notADigit");
        expect(() => typeBelowLastRow(merged, flow.id)).not.toThrow();
    });
});

describe("a stamp no clock produced", () => {
    it("loses the merge instead of winning every one of them", () => {
        const round = roundWithData();
        const flow = round.sheets.find((s) => s.kind !== "cx")!;
        const local = seedDoc(round);
        const ctx: OpContext = { actor: "me", clock: createClock("me", () => 5_000) };
        const mine = applyOp(
            local,
            { kind: "cellText", sheetId: flow.id, col: 0, row: 0, text: "mine" },
            ctx,
        );
        const target = Object.values(mine.sheets[flow.id].cells).find(
            (c) => c.col === 0 && c.text === "mine",
        )!;

        // JSON has no NaN literal. 1e400 parses to Infinity, and
        // Infinity - Infinity is NaN, which compares false against zero and so
        // reads as "the other side is greater" at every call site.
        const incoming = inboundDoc(
            JSON.stringify({
                type: "delta",
                doc: {
                    roundId: round.id,
                    round: {},
                    sheets: {
                        [flow.id]: {
                            id: flow.id,
                            fields: {},
                            deleted: null,
                            cells: {
                                [cellKey(0, target.rank, target.actor)]: hostileCell({
                                    col: 0,
                                    rank: target.rank,
                                    actor: target.actor,
                                    textStamp: { ms: 1e400, counter: 0, actor: "attacker" },
                                }),
                            },
                        },
                    },
                },
            }),
        );

        const merged = merge(mine, incoming).doc;
        const settled = merged.sheets[flow.id].cells[cellKey(0, target.rank, target.actor)];
        expect(settled.text).toBe("mine");
    });

    it("is refused outright in a vector, where every value is checked", () => {
        const stamped = (ms: unknown, counter: unknown) =>
            onWire(
                JSON.stringify({ type: "vector", seen: { x: { ms, counter, actor: "attacker" } } }),
            );
        expect(stamped(1e400, 0)).toBeNull();
        expect(stamped(-1, 0)).toBeNull();
        expect(stamped(1.5, 0)).toBeNull();
        expect(stamped(9_007_199_254_740_992, 0)).toBeNull();
        expect(stamped(1, 9_007_199_254_740_992)).toBeNull();
        expect(stamped("9", 0)).toBeNull();
        expect(stamped(1, 2)).not.toBeNull();
    });
});

describe("what a peer can make this replica retain", () => {
    it("stops taking on new cells at the ceiling and keeps the ones it holds", () => {
        const round = makeFlowRound({});
        const flow = round.sheets.find((s) => s.kind !== "cx")!;
        flow.data = [];
        const local = seedDoc(round);

        const cells: Record<string, CollabCell> = {};
        for (let i = 0; i < 200_100; i++) {
            // Never a trailing zero digit, so every one of these is orderable
            // and the ceiling is the only thing refusing them.
            cells[`k${i}`] = hostileCell({ rank: `${i}1` });
        }
        const flood: CollabDoc = {
            roundId: round.id,
            round: {},
            sheets: { [flow.id]: { id: flow.id, fields: {}, deleted: null, cells } },
        };

        const once = merge(local, flood).doc;
        expect(Object.keys(once.sheets[flow.id].cells)).toHaveLength(200_000);
        // A second flood of fresh keys buys nothing, and nothing already held
        // is evicted, so the two sides still agree on every cell they share.
        const twice = merge(once, flood).doc;
        expect(Object.keys(twice.sheets[flow.id].cells)).toHaveLength(200_000);
        expect(twice.sheets[flow.id].cells.k0).toEqual(once.sheets[flow.id].cells.k0);
    });

    it("stops taking on new register paths and new sheets at their ceilings", () => {
        const round = makeFlowRound({});
        const local = seedDoc(round);

        const registers: CollabDoc["round"] = {};
        for (let i = 0; i < 5_000; i++) registers[`junk${i}`] = { value: "x", stamp: HIGH };
        const sheets: CollabDoc["sheets"] = {};
        for (let i = 0; i < 600; i++) {
            sheets[`sheet${i}`] = { id: `sheet${i}`, fields: {}, deleted: null, cells: {} };
        }

        const merged = merge(local, { roundId: round.id, round: registers, sheets }).doc;
        expect(Object.keys(merged.round)).toHaveLength(4_096);
        expect(Object.keys(merged.sheets)).toHaveLength(512);
    });
});

describe("a prototype key from a peer", () => {
    it("is refused before it reaches a bracket assignment", () => {
        const doc = (round: string, sheets: string) =>
            `{"type":"delta","doc":{"roundId":"r","round":${round},"sheets":${sheets}}}`;
        const sheet = (cells: string, fields = "{}") =>
            `{"s":{"id":"s","fields":${fields},"deleted":null,"cells":${cells}}}`;

        expect(onWire(doc('{"__proto__":{}}', "{}"))).toBeNull();
        expect(onWire(doc("{}", '{"__proto__":{}}'))).toBeNull();
        expect(onWire(doc("{}", sheet('{"__proto__":{}}')))).toBeNull();
        expect(onWire(doc("{}", sheet("{}", '{"__proto__":{}}')))).toBeNull();
        expect(onWire(doc("{}", sheet("{}")))).not.toBeNull();
        expect(({} as Record<string, unknown>).polluted).toBeUndefined();
    });

    it("cannot become a vector key by way of a stamp's actor", () => {
        const round = makeFlowRound({});
        const doc = seedDoc(round);
        doc.round.event = { value: "pf", stamp: { ms: 5, counter: 0, actor: "__proto__" } };
        expect(Object.getPrototypeOf(vectorOf(doc))).toBe(Object.prototype);
    });
});

describe("a sheet named by a key that disagrees with its own id", () => {
    it("projects once, under the key it arrived on", () => {
        const round = roundWithData();
        const flow = round.sheets.find((s) => s.kind !== "cx")!;
        const incoming = deltaOf({
            roundId: round.id,
            round: {},
            // A second entry claiming the existing sheet's id would otherwise
            // project two sheets under one id, and every later local edit to it
            // would be looked up by the key and land nowhere.
            sheets: {
                impostor: {
                    id: flow.id,
                    fields: { title: { value: "mine now", stamp: HIGH } },
                    deleted: null,
                    cells: {},
                },
            },
        });

        const projected = projectDoc(merge(seedDoc(round), incoming).doc, round);
        expect(projected.sheets.filter((s) => s.id === flow.id)).toHaveLength(1);
        expect(projected.sheets.map((s) => s.id)).toContain("impostor");
    });
});
