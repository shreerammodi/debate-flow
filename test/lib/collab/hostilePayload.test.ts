/**
 * What an admitted peer can do with the bytes it is allowed to send.
 *
 * Every case here goes in through `parseWireMessage`, which is the only runtime
 * validator on the peer path, so a test that passes proves the shipping
 * transport would behave the same way. The adversary is a peer that has been
 * admitted once and runs a modified client: it chooses every byte below.
 */

import { describe, expect, it } from "vitest";

import { deltaSince, isEmptyDelta, vectorOf } from "@/lib/collab/delta";
import { projectDoc, seedDoc, sheetWidth } from "@/lib/collab/doc";
import { merge } from "@/lib/collab/merge";
import { applyOp, type OpContext } from "@/lib/collab/ops";
import { parseWireMessage, type WireMessage } from "@/lib/collab/peerLink";
import { createClock, type Stamp } from "@/lib/collab/stamp";
import { cellKey, type CollabCell, type CollabDoc, type CollabSheet } from "@/lib/collab/types";
import { columnsForFlowSheet } from "@/lib/grid/flowColumns";
import { makeFlowRound, type FlowRound } from "@/lib/model/flow";
import { parseFlowFile, serializeFlow } from "@/lib/persistence/flowFile";

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

        const flood = (count: number): CollabDoc => {
            const cells: Record<string, CollabCell> = {};
            for (let i = 0; i < count; i++) {
                // Never a trailing zero digit, so every one of these is
                // orderable and the ceiling is the only thing refusing them.
                cells[`k${i}`] = hostileCell({ rank: `${i}1` });
            }
            return {
                roundId: round.id,
                round: {},
                sheets: { [flow.id]: { id: flow.id, fields: {}, deleted: null, cells } },
            };
        };
        // Two hundred thousand cells as JSON text is gratuitous, so the volume
        // below is a literal. This is what proves the shape of it is one the
        // transport admits, i.e. that the ceiling is reachable over the wire.
        expect(deltaOf(flood(2))).not.toBeNull();

        const once = merge(local, flood(200_100)).doc;
        expect(Object.keys(once.sheets[flow.id].cells)).toHaveLength(200_000);
        // A second flood of fresh keys buys nothing, and nothing already held
        // is evicted, so the two sides still agree on every cell they share.
        const twice = merge(once, flood(200_100)).doc;
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

        // Small enough to send as the bytes a peer would send, so the ceilings
        // are proven to be what refuses this rather than the validator.
        const merged = merge(local, deltaOf({ roundId: round.id, round: registers, sheets })).doc;
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
        // The write those refusals stand in front of: the merge assigns into a
        // plain object literal by the key the message chose, and this asserts
        // the map it hands back is still an ordinary record.
        const register = '{"event":{"value":"pf","stamp":{"ms":1,"counter":0,"actor":"attacker"}}}';
        const admitted = inboundDoc(doc(register, sheet("{}")));
        const merged = merge(seedDoc(makeFlowRound({})), admitted).doc;
        expect(Object.getPrototypeOf(merged.round)).toBe(Object.prototype);
        expect(Object.getPrototypeOf(merged.sheets.s.cells)).toBe(Object.prototype);
    });

    it("is recorded in the vector under its own key rather than read off the chain", () => {
        const round = roundWithData();
        const flow = round.sheets.find((s) => s.kind !== "cx")!;
        const doc = deltaOf({
            roundId: round.id,
            round: {},
            sheets: {
                [flow.id]: {
                    id: flow.id,
                    fields: {},
                    deleted: null,
                    cells: {
                        one: hostileCell({
                            rank: "21",
                            actor: "__proto__",
                            textStamp: { ...HIGH, actor: "__proto__" },
                            metaStamp: { ...HIGH, actor: "__proto__" },
                        }),
                    },
                },
            },
        });

        const seen = vectorOf(doc);
        expect(Object.hasOwn(seen, "__proto__")).toBe(true);
        // An actor no vector can hold is one every 30 ms push and every repair
        // reply re-ships, to every peer, for as long as the round is open.
        expect(isEmptyDelta(deltaSince(doc, seen))).toBe(true);
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

describe("what a peer's registers can do to the file the round is saved as", () => {
    /** The round the store would hold, then the round a reopen reads back. */
    function savedAndReopened(round: FlowRound, doc: unknown): FlowRound {
        const projected = projectDoc(merge(seedDoc(round), deltaOf(doc)).doc, round);
        return parseFlowFile(serializeFlow(projected));
    }

    it("leaves the round openable when the event register names no event", () => {
        const round = roundWithData();
        const reopened = savedAndReopened(round, {
            roundId: round.id,
            round: {
                event: { value: "constructor", stamp: HIGH },
                firstSide: { value: "nobody", stamp: HIGH },
            },
            sheets: {},
        });
        expect(reopened.event).toBe("policy");
        expect(reopened.firstSide).toBe("aff");
    });

    it("leaves it openable when a sheet's own registers are not what a sheet holds", () => {
        const round = roundWithData();
        const flow = round.sheets.find((s) => s.kind !== "cx")!;
        const reopened = savedAndReopened(round, {
            roundId: round.id,
            round: {},
            sheets: {
                [flow.id]: {
                    id: flow.id,
                    deleted: null,
                    cells: {},
                    fields: {
                        title: { value: 7, stamp: HIGH },
                        group: { value: "zzz", stamp: HIGH },
                        order: { value: "first", stamp: HIGH },
                        kind: { value: "whatever", stamp: HIGH },
                        startSpeechId: { value: { nested: true }, stamp: HIGH },
                    },
                },
            },
        });
        const sheet = reopened.sheets.find((s) => s.id === flow.id)!;
        expect(sheet.group).toBe("aff");
        expect(sheet.title).toBe("");
        expect(sheet.order).toBe(0);
        expect(sheet.startSpeechId).toBeUndefined();
    });

    it("leaves it openable when the scouting register is not scouting", () => {
        const round = roundWithData();
        const reopened = savedAndReopened(round, {
            roundId: round.id,
            round: { "scouting.decision.vote": { value: { neither: true }, stamp: HIGH } },
            sheets: {},
        });
        expect(reopened.scouting.decision).toBeUndefined();
    });

    it("leaves it openable when a cell's text and decoration are not either", () => {
        const round = roundWithData();
        const flow = round.sheets.find((s) => s.kind !== "cx")!;
        const reopened = savedAndReopened(round, {
            roundId: round.id,
            round: {},
            sheets: {
                [flow.id]: {
                    id: flow.id,
                    fields: {},
                    deleted: null,
                    cells: {
                        num: hostileCell({ col: 0, rank: "21", text: 5 as unknown as string }),
                        deco: hostileCell({
                            col: 1,
                            rank: "21",
                            meta: { bold: "yes" as unknown as boolean },
                        }),
                    },
                },
            },
        });
        const sheet = reopened.sheets.find((s) => s.id === flow.id)!;
        expect(sheet.data[2]).toEqual([null, "theirs"]);
        expect(sheet.meta["2,1"]).toBeUndefined();
    });

    it("refuses the write rather than putting a round on disk that cannot be reopened", () => {
        const round = roundWithData();
        // The last line of defence, reached only by a bug above it: no write
        // path may produce a file the parser would refuse on the next open.
        const poisoned = { ...round, event: "constructor" as FlowRound["event"] };
        expect(() => serializeFlow(poisoned)).toThrow(/round\.event is not a known debate event/);
    });
});

describe("how tall a peer can make one column", () => {
    it("is bounded where the width is, because the projection allocates the product", () => {
        const round = roundWithData();
        const flow = round.sheets.find((s) => s.kind !== "cx")!;
        const cells: Record<string, CollabCell> = {};
        // One column of four thousand rows plus one cell at the far column: the
        // width bound alone leaves a 4000 x 512 rectangle to materialize.
        for (let i = 0; i < 4_000; i++) cells[`r${i}`] = hostileCell({ col: 0, rank: `${i}1` });
        cells.far = hostileCell({ col: 511, rank: "21" });

        const merged = merge(seedDoc(round), {
            roundId: round.id,
            round: {},
            sheets: { [flow.id]: { id: flow.id, fields: {}, deleted: null, cells } },
        }).doc;
        const projected = projectDoc(merged, round);
        const sheet = projected.sheets.find((s) => s.id === flow.id)!;
        expect(sheet.data).toHaveLength(2_048);
        expect(sheet.data[0]).toHaveLength(512);
        // And the file the autosave writes is still one the app reads back.
        expect(() => parseFlowFile(serializeFlow(projected))).not.toThrow();
    });
});

describe("how many cells a peer can spread across sheets", () => {
    /** One sheet at the widest and tallest this build projects: 2048 x 512. */
    function balloon(id: string): CollabSheet {
        const cells: Record<string, CollabCell> = {};
        for (let i = 0; i < 2_048; i++) cells[`r${i}`] = hostileCell({ col: 0, rank: `${i}1` });
        cells.far = hostileCell({ col: 511, rank: "21" });
        return { id, fields: {}, deleted: null, cells };
    }

    it("cannot put the round past what the file holds, so the autosave still writes", () => {
        const round = roundWithData();
        const flow = round.sheets.find((s) => s.kind !== "cx")!;
        // A real flow for the peer to squeeze: a few hundred rows of eight
        // speeches, which is what the clamp must never reach.
        const typed = Array.from({ length: 220 }, (_, r) =>
            Array.from({ length: 8 }, (_, c) => `arg ${r}.${c}`),
        );
        flow.data = typed;
        // Two of these project to 2,097,152 padded cells against the file's
        // 2,000,000, and the merge's own ceiling is per sheet, so it admits
        // both. Unbudgeted, the projection then produced a round every later
        // write of this flow refused: the peer, not the debater, decided the
        // round could no longer be saved.
        const merged = merge(seedDoc(round), {
            roundId: round.id,
            round: {},
            sheets: { "balloon-a": balloon("balloon-a"), "balloon-b": balloon("balloon-b") },
        }).doc;
        const projected = projectDoc(merged, round);

        const text = serializeFlow(projected);
        expect(parseFlowFile(text)).toEqual(projected);
        expect(serializeFlow(parseFlowFile(text))).toBe(text);

        const reopened = parseFlowFile(text);
        for (const id of ["balloon-a", "balloon-b"]) {
            // Clamped, not dropped: the sheet is still in the round at its full
            // width, and the replica still holds every cell it was sent.
            const sheet = reopened.sheets.find((s) => s.id === id)!;
            expect(sheet.data[0]).toHaveLength(512);
            expect(sheet.data.length).toBeGreaterThan(1_000);
            expect(Object.keys(merged.sheets[id].cells)).toHaveLength(2_049);
        }
        // The debater's own sheet is projected as if the peer had sent nothing,
        // because the budget is spent on the cheapest sheet first and a real
        // sheet is three orders of magnitude cheaper than a ballooned one.
        expect(reopened.sheets.find((s) => s.id === flow.id)!.data).toEqual(typed);
    });
});

describe("a stamp from a peer whose count no clock reported", () => {
    it("does not pin the cell out of the debater's reach", () => {
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
        const key = cellKey(0, target.rank, target.actor);

        // A wall clock at the far end of the safe range, with a count a million
        // times past what a stalled clock could reach.
        const pinned = { ms: 9_007_199_254_740_991, counter: 2_000_000, actor: "attacker" };
        const incoming = deltaOf({
            roundId: round.id,
            round: {},
            sheets: {
                [flow.id]: {
                    id: flow.id,
                    fields: {},
                    deleted: null,
                    cells: {
                        [key]: hostileCell({
                            col: 0,
                            rank: target.rank,
                            actor: target.actor,
                            textStamp: pinned,
                            metaStamp: pinned,
                        }),
                    },
                },
            },
        });

        const merged = merge(mine, incoming).doc;
        // The clock raises off what arrived, which is what every apply does.
        for (const stamp of Object.values(vectorOf(merged))) ctx.clock.observe(stamp);
        const retyped = applyOp(
            merged,
            { kind: "cellText", sheetId: flow.id, col: 0, row: 0, text: "corrected" },
            ctx,
        );
        expect(retyped.sheets[flow.id].cells[key].text).toBe("corrected");

        // And the peer repairing with the same stamp does not take it back.
        expect(merge(retyped, incoming).doc.sheets[flow.id].cells[key].text).toBe("corrected");
    });
});

describe("a document map whose values are not what the merge reads", () => {
    it("is refused at the edge, not thrown out of the middle of a merge", () => {
        const line = (round: string, sheets: string) =>
            `{"type":"delta","doc":{"roundId":"r","round":${round},"sheets":${sheets}}}`;
        const sheet = (cells: string, fields = "{}") =>
            `{"s":{"id":"s","fields":${fields},"deleted":null,"cells":${cells}}}`;
        const stamp = '{"ms":1,"counter":0,"actor":"attacker"}';

        expect(onWire(line("{}", sheet('{"k":null}')))).toBeNull();
        expect(onWire(line('{"event":null}', "{}"))).toBeNull();
        expect(onWire(line('{"event":{"value":"pf"}}', "{}"))).toBeNull();
        expect(onWire(line("{}", sheet("{}", '{"title":{"value":"x"}}')))).toBeNull();
        expect(onWire(line("{}", sheet(`{"k":{"col":0,"textStamp":${stamp}}}`)))).toBeNull();
        // A register whose value is a nested object is ordinary, so the check
        // reaches the stamp and stops there.
        expect(
            onWire(line(`{"scouting":{"value":{"judge":"x"},"stamp":${stamp}}}`, sheet("{}"))),
        ).not.toBeNull();
    });
});
