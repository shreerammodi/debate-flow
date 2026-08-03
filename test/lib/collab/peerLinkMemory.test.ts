import { describe, expect, it } from "vitest";

import type { PeerConn, PeerLinkConfig, WireMessage } from "@/lib/collab/peerLink";
import { createMemoryNet, memoryPairId, memoryRelay } from "@/lib/collab/peerLinkMemory";

const CONFIG: PeerLinkConfig = { discovery: "mdns", relay: true };

describe("peerLinkMemory", () => {
    it("reports the endpoint id it was created for", async () => {
        const net = createMemoryNet();
        const link = await net.create("alex")(CONFIG);
        expect(await link.endpointId()).toBe("alex");
    });

    it("delivers a dial to the listening peer", async () => {
        const net = createMemoryNet();
        const host = await net.create("alex")(CONFIG);
        const guest = await net.create("sam")(CONFIG);
        const arrived: PeerConn[] = [];
        await host.listen((peer) => arrived.push(peer));
        const conn = await guest.dial("alex");
        expect(arrived).toHaveLength(1);
        expect(arrived[0].id).toBe("sam");
        expect(conn.id).toBe("alex");
    });

    it("carries a message both ways", async () => {
        const net = createMemoryNet();
        const host = await net.create("alex")(CONFIG);
        const guest = await net.create("sam")(CONFIG);
        const heard: WireMessage[] = [];
        let hostSide: PeerConn | null = null;
        await host.listen((peer) => {
            hostSide = peer;
            peer.onMessage((m) => heard.push(m));
        });
        const guestSide = await guest.dial("alex");
        const replies: WireMessage[] = [];
        guestSide.onMessage((m) => replies.push(m));

        guestSide.send({ type: "bye" });
        hostSide!.send({ type: "helloAck", ok: true });
        expect(heard).toEqual([{ type: "bye" }]);
        expect(replies).toEqual([{ type: "helloAck", ok: true }]);
    });

    // Load-bearing for every hostile-payload test in the suite: they assert
    // what the far side did with a line, and that only means something if this
    // net narrows a line the way the shipping adapter does. Without this,
    // deleting the narrowing leaves the whole suite green.
    it("drops a line that does not conform to its variant, as the real link does", async () => {
        const net = createMemoryNet();
        const host = await net.create("alex")(CONFIG);
        const guest = await net.create("sam")(CONFIG);
        const heard: WireMessage[] = [];
        await host.listen((peer) => peer.onMessage((m) => heard.push(m)));
        const guestSide = await guest.dial("alex");

        // Parsed rather than written as a literal: `__proto__` in an object
        // literal is the prototype setter, and a peer sends an own key.
        guestSide.send({
            type: "delta",
            doc: { roundId: "r", round: JSON.parse('{"__proto__":{}}'), sheets: {} },
        });
        guestSide.send({ type: "state" } as unknown as WireMessage);
        expect(heard).toEqual([]);

        guestSide.send({ type: "delta", doc: { roundId: "r", round: {}, sheets: {} } });
        expect(heard).toHaveLength(1);
    });

    it("tells both sides when one closes", async () => {
        const net = createMemoryNet();
        const host = await net.create("alex")(CONFIG);
        const guest = await net.create("sam")(CONFIG);
        let hostClosed = false;
        let guestClosed = false;
        await host.listen((peer) => peer.onClose(() => (hostClosed = true)));
        const guestSide = await guest.dial("alex");
        guestSide.onClose(() => (guestClosed = true));
        guestSide.close();
        expect(hostClosed).toBe(true);
        expect(guestClosed).toBe(true);
    });

    it("refuses a dial to an endpoint that is not listening", async () => {
        const net = createMemoryNet();
        const guest = await net.create("sam")(CONFIG);
        await expect(guest.dial("nobody")).rejects.toThrow(/nobody/);
    });

    it("reports a direct connection when neither side allows a relay", async () => {
        const net = createMemoryNet();
        const host = await net.create("alex")({ discovery: "mdns", relay: false });
        const guest = await net.create("sam")({ discovery: "mdns", relay: false });
        await host.listen(() => {});
        expect((await guest.dial("alex")).connectionType()).toBe("direct");
    });

    it("records every call, which is how the off state is proven", async () => {
        const net = createMemoryNet();
        const link = await net.create("alex")(CONFIG);
        await link.listen(() => {});
        await link.stop();
        expect(net.calls.map((c) => c.op)).toEqual(["create", "listen", "stop"]);
        expect(net.calls[0].config).toEqual(CONFIG);
    });

    it("names the endpoint a dial reached out to, and where it looked", async () => {
        const net = createMemoryNet();
        const guest = await net.create("sam")(CONFIG);
        await expect(guest.dial("alex")).rejects.toThrow();
        await expect(guest.dial("kim", "https://relay.example/1")).rejects.toThrow();
        expect(net.calls.filter((c) => c.op === "dial")).toEqual([
            { op: "dial", endpointId: "alex", relayUrl: null },
            { op: "dial", endpointId: "kim", relayUrl: "https://relay.example/1" },
        ]);
    });

    it("is homed on a relay exactly when it would use one", async () => {
        const net = createMemoryNet();
        const relaying = await net.create("alex")({ discovery: "mdns", relay: true });
        const direct = await net.create("sam")({ discovery: "mdns", relay: false });
        expect(await relaying.relayUrl()).toBe(memoryRelay("alex"));
        expect(await direct.relayUrl()).toBe("");
    });

    it("reports where a relayed peer is homed, and nothing for a direct one", async () => {
        const net = createMemoryNet();
        const host = await net.create("alex")({ discovery: "mdns", relay: true });
        const guest = await net.create("sam")({ discovery: "mdns", relay: true });
        let inbound: PeerConn | null = null;
        await host.listen((c) => (inbound = c));

        const out = await guest.dial("alex");
        expect(out.relayUrl()).toBe(memoryRelay("alex"));
        expect(inbound!.relayUrl()).toBe(memoryRelay("sam"));

        const near = await net.create("kim")({ discovery: "mdns", relay: false });
        expect((await near.dial("alex")).relayUrl()).toBeNull();
    });
});

describe("pairing over the memory net", () => {
    it("puts the host and the guest on the same name from one code", async () => {
        const net = createMemoryNet();
        const host = await net.create("alex")(CONFIG);
        const guest = await net.create("sam")(CONFIG);
        const arrived: string[] = [];
        const id = await host.pairHost("TESTAA01", (conn) => arrived.push(conn.id));
        expect(id).toBe(memoryPairId("TESTAA01"));

        const conn = await guest.pairDial("testaa01");
        expect(conn.id).toBe(memoryPairId("TESTAA01"));
        expect(arrived).toEqual(["sam"]);
    });

    it("reads a code the same however it is typed", async () => {
        const net = createMemoryNet();
        const host = await net.create("alex")(CONFIG);
        const guest = await net.create("sam")(CONFIG);
        await host.pairHost("TESTAA01", () => {});
        await expect(guest.pairDial("testaa-01")).resolves.toBeDefined();
    });

    it("refuses a dial once the host has let the code go", async () => {
        const net = createMemoryNet();
        const host = await net.create("alex")(CONFIG);
        const guest = await net.create("sam")(CONFIG);
        await host.pairHost("TESTAA01", () => {});
        await host.pairStop();
        await expect(guest.pairDial("TESTAA01")).rejects.toThrow();
    });

    it("holds one code at a time, so a second replaces the first", async () => {
        const net = createMemoryNet();
        const host = await net.create("alex")(CONFIG);
        const guest = await net.create("sam")(CONFIG);
        await host.pairHost("TESTAA01", () => {});
        await host.pairHost("TESTAA02", () => {});
        await expect(guest.pairDial("TESTAA01")).rejects.toThrow();
        await expect(guest.pairDial("TESTAA02")).resolves.toBeDefined();
    });

    it("takes a code off the air when the link stops", async () => {
        const net = createMemoryNet();
        const host = await net.create("alex")(CONFIG);
        const guest = await net.create("sam")(CONFIG);
        await host.pairHost("TESTAA01", () => {});
        await host.stop();
        await expect(guest.pairDial("TESTAA01")).rejects.toThrow();
    });

    it("records every pairing call, so an opt-in test can read an empty log", async () => {
        const net = createMemoryNet();
        const host = await net.create("alex")(CONFIG);
        await host.newCode();
        await host.pairHost("TESTAA01", () => {});
        await host.pairStop();
        expect(net.calls.map((c) => c.op)).toEqual(["create", "newCode", "pairHost", "pairStop"]);
    });

    it("mints a code the shell would accept back", async () => {
        const net = createMemoryNet();
        const host = await net.create("alex")(CONFIG);
        expect(await host.newCode()).toMatch(/^[0-9A-HJKMNP-TV-Z]{8}$/);
    });
});
