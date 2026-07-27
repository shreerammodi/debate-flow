import { beforeEach, describe, expect, it, vi } from "vitest";

import type { PeerConn, WireMessage } from "@/lib/collab/peerLink";
import { createPeerLink, type TauriBridge } from "@/lib/collab/peerLinkTauri";

/** Stands in for the shell: records commands and replays events by hand. */
function fakeBridge() {
    const calls: { cmd: string; args: Record<string, unknown> }[] = [];
    const listeners: Record<string, ((payload: unknown) => void)[]> = {};
    /** Commands the shell should refuse, the way it refuses a dead connection. */
    const refuse = new Set<string>();
    const bridge: TauriBridge = {
        async invoke(cmd, args) {
            calls.push({ cmd, args });
            if (refuse.has(cmd)) throw new Error("That peer is gone");
            if (cmd === "collab_start") return "alex";
            if (cmd === "collab_dial") return { connId: "c1", connectionType: "direct" };
            return undefined;
        },
        async listen(event, cb) {
            (listeners[event] ??= []).push(cb);
            return () => {
                listeners[event] = listeners[event].filter((l) => l !== cb);
            };
        },
    };
    return {
        bridge,
        calls,
        refuse,
        emit(event: string, payload: unknown) {
            for (const cb of listeners[event] ?? []) cb(payload);
        },
    };
}

let fake: ReturnType<typeof fakeBridge>;

beforeEach(() => {
    fake = fakeBridge();
});

describe("createPeerLink", () => {
    it("starts the endpoint with the config it was given", async () => {
        await createPeerLink({ discovery: "mdns", relay: false }, fake.bridge);
        expect(fake.calls[0]).toEqual({
            cmd: "collab_start",
            args: { relay: false, mdns: true },
        });
    });

    it("turns discovery off when the config says off", async () => {
        await createPeerLink({ discovery: "off", relay: true }, fake.bridge);
        expect(fake.calls[0].args).toEqual({ relay: true, mdns: false });
    });

    it("reports the endpoint id the shell bound", async () => {
        const link = await createPeerLink({ discovery: "mdns", relay: true }, fake.bridge);
        expect(await link.endpointId()).toBe("alex");
    });

    it("hands an inbound peer to the listener", async () => {
        const link = await createPeerLink({ discovery: "mdns", relay: true }, fake.bridge);
        const seen: PeerConn[] = [];
        await link.listen((conn) => seen.push(conn));
        fake.emit("collab:peer", {
            connId: "c9",
            endpointId: "sam",
            connectionType: "relayed",
        });
        expect(seen).toHaveLength(1);
        expect(seen[0].id).toBe("sam");
        expect(seen[0].connectionType()).toBe("relayed");
    });

    it("discloses anything the shell did not call direct as relayed", async () => {
        const link = await createPeerLink({ discovery: "mdns", relay: true }, fake.bridge);
        const seen: PeerConn[] = [];
        await link.listen((conn) => seen.push(conn));
        fake.emit("collab:peer", { connId: "c9", endpointId: "sam" });
        expect(seen[0].connectionType()).toBe("relayed");
    });

    it("does not hand a dialled peer to the inbound listener", async () => {
        const link = await createPeerLink({ discovery: "mdns", relay: true }, fake.bridge);
        const seen: PeerConn[] = [];
        await link.listen((conn) => seen.push(conn));
        await link.dial("sam");
        // The dial reports its own path, so no event ever correlates with it.
        fake.emit("collab:peer", { connId: "c1", endpointId: "sam", connectionType: "direct" });
        expect(seen).toHaveLength(0);
    });

    it("reports the path the dial itself returned", async () => {
        const link = await createPeerLink({ discovery: "mdns", relay: true }, fake.bridge);
        expect((await link.dial("sam")).connectionType()).toBe("direct");
    });

    it("delivers a message to the connection it belongs to", async () => {
        const link = await createPeerLink({ discovery: "mdns", relay: true }, fake.bridge);
        const conn = await link.dial("sam");
        const heard: WireMessage[] = [];
        conn.onMessage((m) => heard.push(m));
        fake.emit("collab:message", { connId: "c1", payload: JSON.stringify({ type: "bye" }) });
        expect(heard).toEqual([{ type: "bye" }]);
    });

    it("ignores a message for a connection it does not hold", async () => {
        const link = await createPeerLink({ discovery: "mdns", relay: true }, fake.bridge);
        const conn = await link.dial("sam");
        const heard: WireMessage[] = [];
        conn.onMessage((m) => heard.push(m));
        fake.emit("collab:message", { connId: "other", payload: '{"type":"bye"}' });
        expect(heard).toEqual([]);
    });

    it("survives a payload that is not a wire message", async () => {
        const link = await createPeerLink({ discovery: "mdns", relay: true }, fake.bridge);
        const conn = await link.dial("sam");
        const heard: WireMessage[] = [];
        conn.onMessage((m) => heard.push(m));
        expect(() =>
            fake.emit("collab:message", { connId: "c1", payload: "{not json" }),
        ).not.toThrow();
        expect(heard).toEqual([]);
    });

    it("sends a message as one JSON line through the shell", async () => {
        const link = await createPeerLink({ discovery: "mdns", relay: true }, fake.bridge);
        const conn = await link.dial("sam");
        conn.send({ type: "bye" });
        const sent = fake.calls.find((c) => c.cmd === "collab_send");
        expect(sent!.args).toEqual({ connId: "c1", payload: '{"type":"bye"}' });
    });

    it("tells a connection when the shell says it closed", async () => {
        const link = await createPeerLink({ discovery: "mdns", relay: true }, fake.bridge);
        const conn = await link.dial("sam");
        let closed = false;
        conn.onClose(() => (closed = true));
        fake.emit("collab:closed", { connId: "c1" });
        expect(closed).toBe(true);
    });

    it("closes a connection through the shell exactly once", async () => {
        const link = await createPeerLink({ discovery: "mdns", relay: true }, fake.bridge);
        const conn = await link.dial("sam");
        const onClose = vi.fn();
        conn.onClose(onClose);
        conn.close();
        conn.close();
        expect(onClose).toHaveBeenCalledTimes(1);
        expect(fake.calls.filter((c) => c.cmd === "collab_close")).toHaveLength(1);
    });

    it("stops the endpoint and drops its listeners", async () => {
        const link = await createPeerLink({ discovery: "mdns", relay: true }, fake.bridge);
        const seen: PeerConn[] = [];
        await link.listen((conn) => seen.push(conn));
        await link.stop();
        expect(fake.calls.some((c) => c.cmd === "collab_stop")).toBe(true);
        fake.emit("collab:peer", { connId: "c9", endpointId: "kim", connectionType: "direct" });
        expect(seen).toHaveLength(0);
    });

    // The shell refuses a send only for a connection it no longer holds: a peer
    // that quit, or an endpoint that stopped. Both are ordinary, and neither
    // may reach the debater as an unhandled rejection.
    it("drops a connection the shell will not send on, without rejecting", async () => {
        const link = await createPeerLink({ discovery: "mdns", relay: true }, fake.bridge);
        const conn = await link.dial("sam");
        const onClose = vi.fn();
        conn.onClose(onClose);

        fake.refuse.add("collab_send");
        conn.send({ type: "bye" });
        await vi.waitFor(() => expect(onClose).toHaveBeenCalledTimes(1));

        // Gone is gone: the link stops trying, and the shell hears no more.
        const before = fake.calls.filter((c) => c.cmd === "collab_send").length;
        conn.send({ type: "bye" });
        expect(fake.calls.filter((c) => c.cmd === "collab_send")).toHaveLength(before);
    });

    it("reports a refused send once, not once per message", async () => {
        const link = await createPeerLink({ discovery: "mdns", relay: true }, fake.bridge);
        const conn = await link.dial("sam");
        const onClose = vi.fn();
        conn.onClose(onClose);

        fake.refuse.add("collab_send");
        conn.send({ type: "bye" });
        conn.send({ type: "bye" });
        conn.send({ type: "bye" });
        await vi.waitFor(() => expect(onClose).toHaveBeenCalledTimes(1));
        fake.emit("collab:closed", { connId: "c1" });
        expect(onClose).toHaveBeenCalledTimes(1);
    });

    it("survives a close the shell refuses, which a peer that hung up first causes", async () => {
        const link = await createPeerLink({ discovery: "mdns", relay: true }, fake.bridge);
        const conn = await link.dial("sam");
        fake.refuse.add("collab_close");
        expect(() => conn.close()).not.toThrow();
    });
});
