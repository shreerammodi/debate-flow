/**
 * One message that used to empty the debater's own sheet for good.
 *
 * Everything here goes through the shipping path in order - `parseWireMessage`,
 * `merge`, `projectDoc`, `serializeFlow`, `parseFlowFile`, and the `healReplica`
 * the end of a session re-arms - because the destruction was in the seam between
 * them: the projection clamped the sheet to fit the round's budget, and the
 * drift check then read that clamp as the replica having fallen behind its own
 * store copy and re-seeded the replica from it. Either half alone leaves the
 * debater's flow gone.
 */

import { describe, expect, it } from "vitest";

import { projectDoc } from "@/lib/collab/doc";
import { merge } from "@/lib/collab/merge";
import { parseWireMessage } from "@/lib/collab/peerLink";
import { seedRank } from "@/lib/collab/rank";
import {
    clearReplica,
    getReplica,
    healReplica,
    replaceReplicaDoc,
    seedReplica,
    setLocalChangeListener,
} from "@/lib/collab/replica";
import { cellKey, type CollabDoc } from "@/lib/collab/types";
import { makeFlowRound, makeFlowSheet, type FlowRound } from "@/lib/model/flow";
import {
    MAX_ROUND_CELLS,
    paddedCells,
    parseFlowFile,
    serializeFlow,
} from "@/lib/persistence/flowFile";

/** Far above any stamp the debater's own clock has reached, and a safe integer. */
const HIGH = { ms: 9_000_000_000_000, counter: 0, actor: "attacker" };

/** A fat but ordinary elim: six sheets, a few hundred rows, eight speeches. */
function realisticRound(): FlowRound {
    const round = makeFlowRound({});
    while (round.sheets.length < 6) {
        const order = round.sheets.length;
        round.sheets.push(makeFlowSheet({ title: `${order}.`, group: "neg", order }));
    }
    for (const sheet of round.sheets) {
        sheet.data = Array.from({ length: 220 }, (_, r) =>
            Array.from({ length: 8 }, (_, c) => (r % 4 === 0 ? null : `arg ${r}.${c}`)),
        );
    }
    return round;
}

function peerCell(col: number, rank: string, text: string) {
    return {
        col,
        rank,
        actor: "attacker",
        text,
        meta: {},
        textStamp: HIGH,
        metaStamp: HIGH,
        deleted: null,
    };
}

/** Ten cells in the first column, which is what makes a sheet cost anything. */
function cheapColumn(): Record<string, unknown> {
    return Object.fromEntries(
        Array.from({ length: 10 }, (_, r) => [
            cellKey(0, seedRank(r), "attacker"),
            peerCell(0, seedRank(r), "cheap"),
        ]),
    );
}

describe("a peer that widens the debater's own sheet", () => {
    it("takes no rows off it, and leaves the replica whole after the session", () => {
        clearReplica();
        const round = realisticRound();
        const victim = round.sheets.find((s) => s.kind !== "cx")!;
        seedReplica(round, "me");
        // A live session, which is also what stops `healReplica` acting mid-round.
        setLocalChangeListener(() => {});

        // One delta, inside the 4 MiB line the shell reads: a cell at the far
        // column of the debater's own sheet and six values of the fattest text
        // the transport takes in columns past its last speech, plus 505 sheets
        // of eleven cells. None of those is a flow cell, and the cheapest way
        // to a share small enough to matter used to be exactly this.
        const sheets: Record<string, unknown> = {
            [victim.id]: {
                id: victim.id,
                fields: {},
                deleted: null,
                cells: {
                    far: peerCell(511, "zz", "far"),
                    ...Object.fromEntries(
                        Array.from({ length: 6 }, (_, i) => [
                            `fat${i}`,
                            peerCell(8 + i, "a1", "x".repeat(16_000)),
                        ]),
                    ),
                },
            },
        };
        for (let n = 0; n < 505; n++) {
            const id = `aaa-peer-${String(n).padStart(4, "0")}`;
            sheets[id] = {
                id,
                fields: {},
                deleted: null,
                cells: { far: peerCell(511, "zz", "far"), ...cheapColumn() },
            };
        }
        const line = JSON.stringify({
            type: "delta",
            doc: { roundId: round.id, round: {}, sheets },
        });
        expect(line.length).toBeLessThan(4 * 1024 * 1024);

        const message = parseWireMessage(JSON.parse(line));
        expect(message?.type).toBe("delta");
        if (message?.type !== "delta") throw new Error("the transport refused the line");
        const merged = merge(getReplica()!, message.doc);
        replaceReplicaDoc(merged.doc);

        // The file the autosave writes next.
        const projected = projectDoc(merged.doc, round);
        const mine = projected.sheets.find((s) => s.id === victim.id)!;
        expect(mine.data).toHaveLength(220);
        expect(mine.data[0]).toHaveLength(8);
        expect(projected.sheets.reduce((n, s) => n + paddedCells(s.data), 0)).toBeLessThanOrEqual(
            MAX_ROUND_CELLS,
        );
        const reopened = parseFlowFile(serializeFlow(projected));
        expect(reopened.sheets.find((s) => s.id === victim.id)!.data[1][0]).toBe("arg 1.0");

        // The session ends, which re-arms the heal against the file above.
        const before = Object.keys(getReplica()!.sheets[victim.id].cells).length;
        setLocalChangeListener(null);
        expect(healReplica(reopened)).toEqual([]);
        expect(Object.keys(getReplica()!.sheets[victim.id].cells)).toHaveLength(before);

        // And the debater deleting the peer's sheets is not needed to get the
        // rows back, so it cannot fail to.
        const alone: CollabDoc = {
            ...getReplica()!,
            sheets: Object.fromEntries(
                Object.entries(getReplica()!.sheets).filter(([id]) => !id.startsWith("aaa-peer-")),
            ),
        };
        const after = projectDoc(alone, reopened);
        expect(after.sheets.find((s) => s.id === victim.id)!.data).toHaveLength(220);
    }, 30_000);
});
