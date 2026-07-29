/**
 * The three switches, watched for the app's lifetime.
 *
 * The master switch is what the product presents as the kill switch, so what
 * matters here is that off is off for a session already running, not only for
 * the routes into a new one. Relay is watched for the same reason one step
 * down: it is chosen when an endpoint binds, so an idle listener has to rebind
 * to honour a debater withdrawing that consent.
 */

import { render, waitFor } from "@testing-library/react";
import { beforeEach, describe, expect, it, vi } from "vitest";

import type { PeerLink, PeerLinkConfig } from "@/lib/collab/peerLink";
import type { MemoryNet } from "@/lib/collab/peerLinkMemory";

vi.mock("sonner", () => ({
    toast: Object.assign(() => {}, {
        warning: () => {},
        error: () => {},
        success: () => {},
        info: () => {},
    }),
}));

/** Filled in below, once the real transport module has been imported. */
const transport = vi.hoisted(() => ({
    link: null as ((config: PeerLinkConfig) => Promise<PeerLink>) | null,
}));

vi.mock("@/lib/collab/peerLink", async (importOriginal) => ({
    ...(await importOriginal<typeof import("@/lib/collab/peerLink")>()),
    createPeerLinkFor: (config: PeerLinkConfig) => transport.link!(config),
}));

import { createMemoryNet } from "@/lib/collab/peerLinkMemory";
import { clearReplica } from "@/lib/collab/replica";
import { currentSession, endSession, startForRound } from "@/lib/collab/runtime";
import { useInviteWatch } from "@/lib/collab/useInviteWatch";
import { makeFlowRound, type FlowRound } from "@/lib/model/flow";
import { useCollabStore } from "@/lib/store/useCollabStore";
import { useFlowStore } from "@/lib/store/useFlowStore";

const net: MemoryNet = createMemoryNet();
const ME = "e".repeat(64);
transport.link = net.create(ME);

function Watcher(): null {
    useInviteWatch();
    return null;
}

/** Mounts the watcher the root layout mounts, one per window. */
function watch(): { unmount(): void } {
    return render(<Watcher />);
}

let round: FlowRound;

beforeEach(async () => {
    // The watcher is desktop-only, and isDesktop() reads this global.
    (window as unknown as Record<string, unknown>).__TAURI_INTERNALS__ = {};
    await endSession();
    net.reset();
    clearReplica();
    useCollabStore.getState().reset();
    useFlowStore.setState({
        collabEnabled: true,
        collabRelayEnabled: true,
        collabListenEnabled: false,
        contacts: {},
        docPath: "/flows/round-3-harvard.ebb",
    });
    round = makeFlowRound({});
});

describe("throwing the master switch off", () => {
    // A debater realises mid-round that a scout is in the room and reaches for
    // the control the product presents as the kill switch. Every keystroke
    // after that must stop reaching the peer, and the endpoint must go too.
    it("ends a session that is running and lets go of the endpoint", async () => {
        const view = watch();
        expect(await startForRound(round)).not.toBeNull();
        net.reset();

        useFlowStore.setState({ collabEnabled: false });

        await waitFor(() => expect(currentSession()).toBeNull());
        expect(net.calls.filter((c) => c.op === "stop")).toHaveLength(1);
        expect(useCollabStore.getState().status).toBe("off");
        view.unmount();
    });

    it("leaves the session alone while the switch stays on", async () => {
        const view = watch();
        const session = await startForRound(round);

        useFlowStore.setState({ collabRelayEnabled: false });
        await Promise.resolve();

        expect(currentSession()).toBe(session);
        view.unmount();
    });
});

describe("turning Allow relay off", () => {
    // The relay is chosen at bind time, so an idle endpoint bound through one
    // keeps this install visible to a relay operator until it rebinds.
    it("rebinds the idle listener without the relay", async () => {
        useFlowStore.setState({ collabListenEnabled: true });
        const view = watch();
        await waitFor(() => expect(net.calls.filter((c) => c.op === "listen")).toHaveLength(1));

        useFlowStore.setState({ collabRelayEnabled: false });

        await waitFor(() => expect(net.calls.filter((c) => c.op === "listen")).toHaveLength(2));
        expect(net.calls.filter((c) => c.op === "stop")).toHaveLength(1);
        expect(net.calls.filter((c) => c.op === "create").map((c) => c.config!.relay)).toEqual([
            true,
            false,
        ]);
        view.unmount();
    });

    it("binds nothing at all while Listen for invites is off", async () => {
        const view = watch();

        useFlowStore.setState({ collabRelayEnabled: false });

        await Promise.resolve();
        expect(net.calls).toEqual([]);
        view.unmount();
    });
});
