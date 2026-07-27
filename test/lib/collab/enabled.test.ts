import { afterEach, beforeEach, describe, expect, it } from "vitest";

import { collabLive, collabSettings } from "@/lib/collab/enabled";
import { createPeerLinkFor } from "@/lib/collab/peerLink";
import { useFlowStore } from "@/lib/store/useFlowStore";

/**
 * Shared editing is an iroh endpoint, which a browser cannot bind, so it is
 * offered on the desktop and nowhere else. What follows is the boundary: which
 * question each gate answers, and what happens to anything that gets past
 * both.
 */

/** jsdom is not the shell; this is what the shell's presence looks like. */
function pretendDesktop(on: boolean): void {
    const shell = window as unknown as Record<string, unknown>;
    if (on) shell.__TAURI_INTERNALS__ = {};
    else delete shell.__TAURI_INTERNALS__;
}

beforeEach(() => {
    useFlowStore.setState({ collabEnabled: true, collabRelayEnabled: true });
});

afterEach(() => {
    pretendDesktop(false);
});

describe("whether shared editing is offered here", () => {
    it("needs the switch and the shell together", () => {
        pretendDesktop(true);
        expect(collabLive()).toBe(true);

        useFlowStore.setState({ collabEnabled: false });
        expect(collabLive()).toBe(false);
    });

    it("is false on the web however the switch is set", () => {
        pretendDesktop(false);
        expect(collabLive()).toBe(false);
        useFlowStore.setState({ collabEnabled: true });
        expect(collabLive()).toBe(false);
    });
});

describe("what the switch says", () => {
    // Asked by code that has already been handed a transport, so it reports
    // the switch and not the runtime. That is what lets the suite drive the
    // protocol against an in-process transport.
    it("reports the switch alone, so an injected transport still runs", () => {
        pretendDesktop(false);
        expect(collabSettings()).toEqual({ enabled: true, relay: true });

        useFlowStore.setState({ collabEnabled: false, collabRelayEnabled: false });
        expect(collabSettings()).toEqual({ enabled: false, relay: false });
    });
});

describe("the transport a session is given", () => {
    // The backstop. Nothing should reach here off the desktop, and a stand-in
    // that satisfied the port would be worse than the throw: it would mint
    // tickets nobody can redeem and report a peer that cannot exist.
    it("refuses to be resolved anywhere but the desktop", async () => {
        pretendDesktop(false);
        await expect(createPeerLinkFor({ discovery: "mdns", relay: true })).rejects.toThrow(
            /desktop/i,
        );
    });

    it("hands back no in-process stand-in for the browser", async () => {
        pretendDesktop(false);
        await expect(createPeerLinkFor({ discovery: "off", relay: false })).rejects.toBeInstanceOf(
            Error,
        );
    });
});
