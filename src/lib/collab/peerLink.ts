/**
 * The transport ebb's peers talk over.
 *
 * One port, two adapters: iroh in the desktop shell, and an in-process map for
 * the test suite. Everything above this line - the session, presence, the
 * merge - is written against the port, which is what lets convergence be
 * proven without opening a socket. This mirrors the FlowFs port and
 * flowFsMemory.
 */

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
          /** Present only on the first join, and spent when it is accepted. */
          ticket?: string;
      }
    | { type: "helloAck"; ok: true }
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
