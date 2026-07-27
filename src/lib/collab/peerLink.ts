/**
 * The transport ebb's peers talk over.
 *
 * One port, two adapters: iroh in the desktop shell, and an in-process map for
 * the test suite. Everything above this line - the session, presence, the
 * merge - is written against the port, which is what lets convergence be
 * proven without opening a socket. This mirrors the FlowFs port and
 * flowFsMemory.
 */

import { isDesktop } from "@/lib/update/adapter";

import type { Stamp } from "./stamp";
import type { CollabDoc, Role } from "./types";

/** Bumped only for a change an older build cannot read. */
export const PROTOCOL_MAJOR = 1;

export type WireMessage =
    | {
          type: "hello";
          protocol: number;
          app: string;
          endpointId: string;
          roundId: string;
          role: Role;
          capabilities: string[];
          /** What the dialler calls this round, for an invite's corner message. */
          label?: string;
          /**
           * What the dialler calls themselves. A suggestion the far side may
           * show and may save; a name a contact already carries wins over it,
           * because that one is the receiver's own word for this peer.
           */
          name?: string;
          /** Present only on the first join, and spent when it is accepted. */
          ticket?: string;
      }
    /** The host answers with its own name, so naming works in both directions. */
    | { type: "helloAck"; ok: true; name?: string }
    | { type: "helloAck"; ok: false; reason: string }
    | { type: "state"; doc: CollabDoc }
    | { type: "delta"; doc: CollabDoc }
    /** Per-actor highest stamp seen, so the far side can replay what was lost. */
    | { type: "vector"; seen: Record<string, Stamp> }
    | { type: "presence"; cell: { sheetId: string; col: number; row: number } | null }
    | { type: "bye" };

export interface PeerLinkConfig {
    /**
     * DNS discovery is not an option here. An idle ebb publishes nothing about
     * itself anywhere; mDNS reaches the machine across the room and nothing
     * further.
     */
    discovery: "off" | "mdns";
    /** Follows the Allow relay setting. Off restricts a session to direct links. */
    relay: boolean;
}

export interface PeerConn {
    /** The far side's EndpointId. */
    id: string;
    connectionType(): "direct" | "relayed";
    send(msg: WireMessage): void;
    onMessage(cb: (msg: WireMessage) => void): void;
    onClose(cb: () => void): void;
    close(): void;
}

export interface PeerLink {
    endpointId(): Promise<string>;
    listen(onPeer: (peer: PeerConn) => void): Promise<void>;
    dial(endpointId: string, ticket?: string): Promise<PeerConn>;
    stop(): Promise<void>;
}

/** Anything that can hand back a link for a configuration. */
export type PeerLinkFactory = (config: PeerLinkConfig) => Promise<PeerLink>;

/**
 * The transport for this runtime: iroh in the desktop shell, an in-process map
 * everywhere else. Resolved per call rather than cached, because a session
 * binds an endpoint and stopping it must actually release it.
 */
export async function createPeerLinkFor(config: PeerLinkConfig): Promise<PeerLink> {
    // Dynamic on both branches so the browser bundle never pulls in Tauri's JS
    // API, matching how every other desktop touchpoint is gated.
    if (isDesktop()) {
        const mod = await import("./peerLinkTauri");
        return mod.createPeerLink(config);
    }
    const mod = await import("./peerLinkMemory");
    return mod.createMemoryNet().create("browser")(config);
}
