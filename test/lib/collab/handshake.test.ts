import { describe, expect, it } from "vitest";

import { admit, helloFrom, type HostPolicy } from "@/lib/collab/handshake";
import { PROTOCOL_MAJOR, type WireMessage } from "@/lib/collab/peerLink";

const SECRET = "s".repeat(24);

function policy(over: Partial<HostPolicy> = {}): HostPolicy {
    return {
        roundId: "round_x_1",
        appVersion: "0.11.0",
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
        const got = admit(hello({ endpointId: "sam" }), policy({ knownPeers: ["sam"] }), "sam");
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

    it("names both versions when the protocol major differs", () => {
        const got = admit(
            hello({ protocol: PROTOCOL_MAJOR + 1, app: "0.13.0", ticket: SECRET }),
            policy(),
            "guest-1",
        );
        expect(got.ok).toBe(false);
        if (!got.ok) {
            expect(got.silent).toBe(false);
            expect(got.reason).toContain("0.13.0");
            expect(got.reason).toContain("0.11.0");
        }
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
