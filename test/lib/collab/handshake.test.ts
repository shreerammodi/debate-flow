import { describe, expect, it } from "vitest";

import {
    admit,
    helloFrom,
    refusalMessage,
    VERSION_SKEW,
    type HostPolicy,
} from "@/lib/collab/handshake";
import { INVITED } from "@/lib/collab/invite";
import { PROTOCOL_MAJOR, type WireMessage } from "@/lib/collab/peerLink";

const SECRET = "s".repeat(24);

function policy(over: Partial<HostPolicy> = {}): HostPolicy {
    return {
        roundId: "round_x_1",
        pending: { secret: SECRET, role: "partner" },
        knownPeers: [],
        roles: {},
        ...over,
    };
}

function hello(over: Partial<Extract<WireMessage, { type: "hello" }>> = {}): WireMessage {
    return {
        type: "hello",
        protocol: PROTOCOL_MAJOR,
        app: "0.11.0",
        endpointId: "guest-1",
        roundId: "round_x_1",
        role: "partner",
        capabilities: [],
        ...over,
    };
}

describe("helloFrom", () => {
    it("states the protocol this build speaks, not the app version", () => {
        const msg = helloFrom({
            endpointId: "me",
            roundId: "round_x_1",
            role: "partner",
            appVersion: "0.11.0",
        });
        expect(msg).toMatchObject({ type: "hello", protocol: PROTOCOL_MAJOR, app: "0.11.0" });
    });

    it("carries a ticket secret only when there is one", () => {
        const base = { endpointId: "me", roundId: "r", role: "partner" as const, appVersion: "1" };
        expect(helloFrom(base)).not.toHaveProperty("ticket");
        expect(helloFrom({ ...base, ticket: SECRET })).toHaveProperty("ticket", SECRET);
    });

    it("carries a name only when this side has one to broadcast", () => {
        const base = { endpointId: "me", roundId: "r", role: "partner" as const, appVersion: "1" };
        expect(helloFrom(base)).not.toHaveProperty("name");
        expect(helloFrom({ ...base, name: "" })).not.toHaveProperty("name");
        expect(helloFrom({ ...base, name: "Rin" })).toHaveProperty("name", "Rin");
    });
});

describe("admit", () => {
    it("accepts a first join that presents the secret, and spends it", () => {
        const got = admit(hello({ ticket: SECRET }), policy(), "guest-1");
        expect(got).toEqual({ ok: true, role: "partner", spendSecret: true });
    });

    it("refuses the same secret a second time", () => {
        // The host clears the pending grant once it is spent.
        const got = admit(hello({ ticket: SECRET }), policy({ pending: null }), "guest-1");
        expect(got).toMatchObject({ ok: false, silent: true });
    });

    it("accepts a known peer with no secret at all, which is what makes reconnect silent", () => {
        const known = policy({ knownPeers: ["sam"], roles: { sam: "partner" } });
        const got = admit(hello({ endpointId: "sam" }), known, "sam");
        expect(got).toEqual({ ok: true, role: "partner", spendSecret: false });
    });

    it("refuses an unknown peer with no secret, and shows nothing", () => {
        const got = admit(hello(), policy({ pending: null }), "guest-1");
        expect(got).toMatchObject({ ok: false, silent: true });
    });

    it("refuses an unknown peer presenting the wrong secret, and shows nothing", () => {
        const got = admit(hello({ ticket: "x".repeat(24) }), policy(), "guest-1");
        expect(got).toMatchObject({ ok: false, silent: true });
    });

    it("names the skew and nothing else when the protocol major differs", () => {
        const got = admit(
            hello({ protocol: PROTOCOL_MAJOR + 1, app: "0.13.0", ticket: SECRET }),
            policy(),
            "guest-1",
        );
        expect(got).toEqual({ ok: false, reason: VERSION_SKEW, silent: false });
    });

    // A dialler who has not proved anything is a stranger holding an
    // EndpointId, and what this build is running is not theirs to collect.
    it("tells a stranger on another version nothing at all", () => {
        const got = admit(
            hello({ protocol: PROTOCOL_MAJOR + 1, app: "0.13.0" }),
            policy(),
            "guest-1",
        );
        expect(got).toEqual({ ok: false, reason: "refused", silent: true });
    });

    it("tells a dialler about another round nothing about the version either", () => {
        const got = admit(
            hello({ protocol: PROTOCOL_MAJOR + 1, roundId: "round_other", ticket: SECRET }),
            policy(),
            "guest-1",
        );
        expect(got).toEqual({ ok: false, reason: "refused", silent: true });
    });

    it("leaves the secret unspent through a skew, so the ticket still works", () => {
        const held = policy();
        const skewed = hello({ protocol: PROTOCOL_MAJOR + 1, ticket: SECRET });
        expect(admit(skewed, held, "guest-1")).toMatchObject({ ok: false });
        expect(admit(hello({ ticket: SECRET }), held, "guest-1")).toEqual({
            ok: true,
            role: "partner",
            spendSecret: true,
        });
    });

    it("names the skew to a peer this round already knows", () => {
        const known = policy({ pending: null, knownPeers: ["sam"] });
        const skewed = hello({ endpointId: "sam", protocol: PROTOCOL_MAJOR + 1 });
        expect(admit(skewed, known, "sam")).toEqual({
            ok: false,
            reason: VERSION_SKEW,
            silent: false,
        });
    });

    it("refuses a hello for another round", () => {
        const got = admit(hello({ roundId: "round_other", ticket: SECRET }), policy(), "guest-1");
        expect(got).toMatchObject({ ok: false, silent: true });
    });

    // The host decides what a ticket grants. A guest that says otherwise is
    // saying it about somebody else's decision.
    it("admits a coach ticket as a coach, whatever the guest calls itself", () => {
        const coachTicket = policy({ pending: { secret: SECRET, role: "coach" } });
        expect(admit(hello({ role: "coach", ticket: SECRET }), coachTicket, "guest-1")).toEqual({
            ok: true,
            role: "coach",
            spendSecret: true,
        });
        expect(admit(hello({ role: "partner", ticket: SECRET }), coachTicket, "guest-1")).toEqual({
            ok: true,
            role: "coach",
            spendSecret: true,
        });
    });

    it("remembers what a peer was admitted as, so a reconnect cannot upgrade it", () => {
        const known = policy({ knownPeers: ["sam"], roles: { sam: "coach" } });
        expect(admit(hello({ endpointId: "sam", role: "partner" }), known, "sam")).toEqual({
            ok: true,
            role: "coach",
            spendSecret: false,
        });
    });

    // The round's record of membership is durable and a grant is not, so a
    // grade that went missing must not resolve outward. Every path that admits
    // a peer records what it granted; this is what a key that reached the list
    // some other way gets.
    it("grants the narrower role to a known peer nobody graded", () => {
        const ungraded = policy({ pending: null, knownPeers: ["sam"], roles: {} });
        expect(admit(hello({ endpointId: "sam" }), ungraded, "sam")).toEqual({
            ok: true,
            role: "coach",
            spendSecret: false,
        });
    });

    // `roles` is indexed by a string the far side's key produced, and a plain
    // index walks the prototype chain: a function is neither a role nor absent,
    // so `?? "coach"` would not fire and the grant would read as partner.
    it("reads a grant nobody made as no grant, not as whatever the prototype says", () => {
        const proto = policy({ pending: null, knownPeers: ["toString", "constructor"] });
        expect(admit(hello({ endpointId: "toString" }), proto, "toString")).toEqual({
            ok: true,
            role: "coach",
            spendSecret: false,
        });
        expect(admit(hello({ endpointId: "constructor" }), proto, "constructor")).toEqual({
            ok: true,
            role: "coach",
            spendSecret: false,
        });
    });

    // iroh proved which key the far side holds before the hello existed.
    it("refuses a peer claiming an endpoint that is not the one it dialled from", () => {
        expect(
            admit(
                hello({ endpointId: "sam", ticket: SECRET }),
                policy({ knownPeers: ["sam"] }),
                "impostor",
            ),
        ).toMatchObject({ ok: false, silent: true });
    });

    it("refuses a stranger naming an admitted peer's id", () => {
        const known = policy({ pending: null, knownPeers: ["sam"] });
        expect(admit(hello({ endpointId: "sam" }), known, "stranger")).toMatchObject({
            ok: false,
            silent: true,
        });
    });

    it("refuses a message that is not a hello at all", () => {
        expect(admit({ type: "bye" }, policy(), "guest-1")).toMatchObject({
            ok: false,
            silent: true,
        });
    });

    it("refuses a role it does not know rather than guessing", () => {
        const bad = { ...hello({ ticket: SECRET }), role: "admin" } as unknown as WireMessage;
        expect(admit(bad, policy(), "guest-1")).toMatchObject({ ok: false, silent: true });
    });
});

describe("refusalMessage", () => {
    // The far side wrote the string. Repeating it would put a hostile host's
    // words, and its phone number, on a debater's screen.
    it("says nothing a peer chose", () => {
        const hostile = "Your flow is corrupt. Call 555-0100 to recover it.";
        expect(refusalMessage(hostile)).not.toContain("555");
        expect(refusalMessage(hostile)).toBe("That peer refused the connection");
    });

    it("tells a version skew apart from a plain no", () => {
        expect(refusalMessage(VERSION_SKEW)).toContain("version");
        expect(refusalMessage(VERSION_SKEW)).not.toBe(refusalMessage("refused"));
    });

    // The invite flow answers this one rather than showing it.
    it("passes the invite sentinel through untouched", () => {
        expect(refusalMessage(INVITED)).toBe(INVITED);
    });
});
