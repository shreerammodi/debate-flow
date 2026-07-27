import { describe, expect, it } from "vitest";

import { admit, helloFrom, type HostPolicy } from "@/lib/collab/handshake";
import { PROTOCOL_MAJOR, type WireMessage } from "@/lib/collab/peerLink";

const SECRET = "s".repeat(24);

function policy(over: Partial<HostPolicy> = {}): HostPolicy {
    return {
        roundId: "round_x_1",
        appVersion: "0.11.0",
        pendingSecret: SECRET,
        knownPeers: [],
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
        const got = admit(hello({ ticket: SECRET }), policy());
        expect(got).toEqual({ ok: true, role: "partner", spendSecret: true });
    });

    it("refuses the same secret a second time", () => {
        // The host clears pendingSecret once it is spent.
        const got = admit(hello({ ticket: SECRET }), policy({ pendingSecret: null }));
        expect(got).toMatchObject({ ok: false, silent: true });
    });

    it("accepts a known peer with no secret at all, which is what makes reconnect silent", () => {
        const got = admit(hello({ endpointId: "sam" }), policy({ knownPeers: ["sam"] }));
        expect(got).toEqual({ ok: true, role: "partner", spendSecret: false });
    });

    it("refuses an unknown peer with no secret, and shows nothing", () => {
        const got = admit(hello(), policy({ pendingSecret: null }));
        expect(got).toMatchObject({ ok: false, silent: true });
    });

    it("refuses an unknown peer presenting the wrong secret, and shows nothing", () => {
        const got = admit(hello({ ticket: "x".repeat(24) }), policy());
        expect(got).toMatchObject({ ok: false, silent: true });
    });

    it("names both versions when the protocol major differs", () => {
        const got = admit(
            hello({ protocol: PROTOCOL_MAJOR + 1, app: "0.13.0", ticket: SECRET }),
            policy(),
        );
        expect(got.ok).toBe(false);
        if (!got.ok) {
            expect(got.silent).toBe(false);
            expect(got.reason).toContain("0.13.0");
            expect(got.reason).toContain("0.11.0");
        }
    });

    it("refuses a hello for another round", () => {
        const got = admit(hello({ roundId: "round_other", ticket: SECRET }), policy());
        expect(got).toMatchObject({ ok: false, silent: true });
    });

    it("admits a coach read-only", () => {
        const got = admit(hello({ role: "coach", ticket: SECRET }), policy());
        expect(got).toEqual({ ok: true, role: "coach", spendSecret: true });
    });

    it("refuses a message that is not a hello at all", () => {
        expect(admit({ type: "bye" }, policy())).toMatchObject({ ok: false, silent: true });
    });

    it("refuses a role it does not know rather than guessing", () => {
        const bad = { ...hello({ ticket: SECRET }), role: "admin" } as unknown as WireMessage;
        expect(admit(bad, policy())).toMatchObject({ ok: false, silent: true });
    });
});
