import { describe, expect, it } from "vitest";

import { encodeTicket, mintTicket, parseTicket, TICKET_PREFIX } from "@/lib/collab/ticket";

/** What iroh hands back, which is the only thing a ticket may name. */
const HOST = "3f".repeat(32);

const input = { endpointId: HOST, roundId: "round_x_1", role: "partner" as const, relay: true };

/** Encodes a hand-built payload the way encodeTicket does, to test refusals. */
function wrap(payload: unknown): string {
    const json = JSON.stringify(payload);
    const b64 = btoa(json).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
    return TICKET_PREFIX + b64;
}

describe("mintTicket", () => {
    it("carries the host, the round, the role, and the relay stance", () => {
        const t = mintTicket(input, () => "s".repeat(24));
        expect(t).toMatchObject({ ...input, secret: "s".repeat(24) });
    });

    it("mints a different secret each time", () => {
        expect(mintTicket(input).secret).not.toBe(mintTicket(input).secret);
    });

    it("mints a secret long enough to not be guessed", () => {
        expect(mintTicket(input).secret).toMatch(/^[A-Za-z0-9]{24}$/);
    });
});

describe("encodeTicket and parseTicket", () => {
    it("round-trips", () => {
        const t = mintTicket(input);
        expect(parseTicket(encodeTicket(t))).toEqual(t);
    });

    it("names itself, so a user can see what they pasted", () => {
        expect(encodeTicket(mintTicket(input)).startsWith(TICKET_PREFIX)).toBe(true);
    });

    it("survives the whitespace a clipboard adds", () => {
        const text = encodeTicket(mintTicket(input));
        expect(parseTicket(`  ${text}\n`)).not.toBeNull();
    });

    it("refuses anything that is not a ticket", () => {
        expect(parseTicket("")).toBeNull();
        expect(parseTicket("hello")).toBeNull();
        expect(parseTicket(`${TICKET_PREFIX}not base64!!`)).toBeNull();
        expect(parseTicket(wrap([]))).toBeNull();
        expect(parseTicket(wrap(null))).toBeNull();
        expect(parseTicket(wrap("a string"))).toBeNull();
    });

    it("refuses a ticket missing any field it must not guess", () => {
        for (const drop of ["endpointId", "roundId", "role", "secret"]) {
            const t: Record<string, unknown> = { ...mintTicket(input) };
            delete t[drop];
            expect(parseTicket(wrap(t))).toBeNull();
        }
    });

    it("refuses a role it does not know", () => {
        expect(parseTicket(wrap({ ...mintTicket(input), role: "admin" }))).toBeNull();
    });

    it("refuses a secret of the wrong shape", () => {
        expect(parseTicket(wrap({ ...mintTicket(input), secret: "" }))).toBeNull();
        expect(parseTicket(wrap({ ...mintTicket(input), secret: "short" }))).toBeNull();
    });

    it("defaults a missing relay stance to off, the safer answer", () => {
        const t: Record<string, unknown> = { ...mintTicket(input) };
        delete t.relay;
        expect(parseTicket(wrap(t))!.relay).toBe(false);
    });

    it("keeps a coach ticket read-only", () => {
        const t = mintTicket({ ...input, role: "coach" });
        expect(parseTicket(encodeTicket(t))!.role).toBe("coach");
    });
});

describe("the endpoint a ticket names", () => {
    // The next thing that happens to a parsed ticket is a dial, so a ticket
    // that names something iroh could not have issued names an attacker's
    // choice of target rather than a host.
    it("refuses anything that is not an EndpointId", () => {
        for (const endpointId of ["abc123", "", "  ", "x".repeat(64), "3f".repeat(31)]) {
            expect(parseTicket(wrap({ ...mintTicket(input), endpointId }))).toBeNull();
        }
    });

    it("takes the z-base-32 form as readily as the hex one", () => {
        const endpointId = "a2".repeat(26);
        expect(parseTicket(wrap({ ...mintTicket(input), endpointId }))).not.toBeNull();
    });

    it("refuses a round id long enough to be a payload", () => {
        const roundId = "r".repeat(129);
        expect(parseTicket(wrap({ ...mintTicket(input), roundId }))).toBeNull();
    });
});
