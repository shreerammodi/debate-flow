import { describe, expect, it, vi } from "vitest";

import { lossMessage } from "@/lib/collab/lossReport";
import type { DroppedCell } from "@/lib/collab/merge";

const buried = (text: string, writtenBy: string, deletedBy: string): DroppedCell => ({
    sheetId: "s1",
    col: 0,
    rank: "a",
    text,
    writtenBy,
    deletedBy,
});

const contacts = { sam: { name: "Sam" } };

describe("lossMessage", () => {
    it("says nothing when a merge buried nothing", () => {
        expect(lossMessage(contacts, [], "me")).toBeNull();
    });

    it("names the partner whose delete buried one cell, and quotes it", () => {
        const msg = lossMessage(contacts, [buried("perm do both", "me", "sam")], "me");
        expect(msg).toBe('Sam deleted a row over your "perm do both"');
    });

    it("counts rather than listing when several cells went at once", () => {
        const msg = lossMessage(
            contacts,
            [buried("one", "me", "sam"), buried("two", "me", "sam")],
            "me",
        );
        expect(msg).toBe("Sam deleted a row over 2 of your cells");
    });

    it("stays quiet about text this machine did not write", () => {
        // A partner's own write buried by another partner is their business,
        // and reporting it would interrupt a debater over someone else's cell.
        expect(lossMessage(contacts, [buried("theirs", "kim", "sam")], "me")).toBeNull();
    });

    it("counts text that came in from the file as the user's own", () => {
        // A seeded cell carries no author. It was on this grid before the
        // session, so a partner deleting the row is this user's loss.
        const msg = lossMessage(contacts, [buried("was in the file", "", "sam")], "me");
        expect(msg).toBe('Sam deleted a row over your \"was in the file\"');
    });

    it("reports only this machine's cells when a merge buried a mix", () => {
        const msg = lossMessage(
            contacts,
            [buried("theirs", "kim", "sam"), buried("mine", "me", "sam")],
            "me",
        );
        expect(msg).toBe('Sam deleted a row over your "mine"');
    });

    it("falls back to a short id for a peer who is not a contact", () => {
        const msg = lossMessage({}, [buried("gone", "me", "k51qzi5uqu5dlstranger")], "me");
        expect(msg).toBe('k51qzi5u deleted a row over your "gone"');
    });

    it("shortens a long cell so a corner message stays one line", () => {
        const long = "a".repeat(80);
        const msg = lossMessage(contacts, [buried(long, "me", "sam")], "me");
        expect(msg).toContain("...");
        expect(msg!.length).toBeLessThan(80);
    });
});

describe("the live apply path", () => {
    it("tells the user when a partner's delete buried their write", async () => {
        const warnings: string[] = [];
        vi.doMock("sonner", () => ({
            toast: Object.assign(() => {}, {
                warning: (m: string) => warnings.push(m),
                error: () => {},
                success: () => {},
            }),
        }));
        vi.resetModules();

        const { seedReplica, getReplica, replicaActor } = await import("@/lib/collab/replica");
        const { applyOp } = await import("@/lib/collab/ops");
        const { merge } = await import("@/lib/collab/merge");
        const { lossMessage: live } = await import("@/lib/collab/lossReport");
        const { makeFlowRound } = await import("@/lib/model/flow");
        const { createClock } = await import("@/lib/collab/stamp");

        const round = makeFlowRound({});
        const sheet = round.sheets.find((s) => s.kind !== "cx")!;
        sheet.data = [["perm do both"], ["turn"]];
        seedReplica(round, "me");

        // The partner deletes the row this machine just wrote into.
        let t = 9_000;
        const theirs = applyOp(
            getReplica()!,
            { kind: "removeRow", sheetId: sheet.id, row: 0 },
            { actor: "sam", clock: createClock("sam", () => t++) },
        );
        const result = merge(getReplica()!, theirs);

        expect(result.dropped.length).toBeGreaterThan(0);
        const msg = live({ sam: { name: "Sam" } }, result.dropped, replicaActor());
        expect(msg).toContain("Sam deleted a row over your");
        expect(msg).toContain("perm do both");
    });
});
