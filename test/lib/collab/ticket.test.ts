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

    it("carries where the host can be found, so a guest elsewhere can reach it", () => {
        const t = mintTicket({ ...input, relayUrl: "https://usw1-1.relay.n0.iroh.link./" });
        expect(parseTicket(encodeTicket(t))!.relayUrl).toBe("https://usw1-1.relay.n0.iroh.link./");
    });

    // A host running with relaying off has no relay to name, and a blank
    // field in a pasted ticket reads as a broken one.
    it("names no relay at all when the host has none", () => {
        expect("relayUrl" in mintTicket({ ...input, relayUrl: "" })).toBe(false);
        expect(parseTicket(encodeTicket(mintTicket(input)))!.relayUrl).toBeUndefined();
    });

    // The relay is a dial target the moment this returns, so a scheme
    // somebody chose is dropped rather than followed. A ticket without one
    // still opens the round for a guest in the same room.
    it("drops a relay that is not an https address, and keeps the ticket", () => {
        for (const relayUrl of [
            "http://relay.example/",
            "file:///etc/passwd",
            "javascript:alert(1)",
            `https://relay.example/${"x".repeat(300)}`,
            42,
        ]) {
            const parsed = parseTicket(wrap({ ...mintTicket(input), relayUrl }));
            expect(parsed).not.toBeNull();
            expect(parsed!.relayUrl).toBeUndefined();
        }
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
