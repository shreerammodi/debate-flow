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

/** A grid slot a peer is on. */
export interface CellRef {
    sheetId: string;
    col: number;
    row: number;
}

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
    /**
     * The host answers with its own name, so naming works in both directions,
     * and with the role it granted, which is the only way the guest learns it:
     * a guest asks to be a partner and is admitted as whatever the ticket that
     * let it in says. An ack with no role is an older host, and partner is
     * what every one of those granted.
     */
    | { type: "helloAck"; ok: true; name?: string; role?: Role }
    | { type: "helloAck"; ok: false; reason: string }
    | { type: "state"; doc: CollabDoc }
    | { type: "delta"; doc: CollabDoc }
    /** Per-actor highest stamp seen, so the far side can replay what was lost. */
    | { type: "vector"; seen: Record<string, Stamp> }
    /** The cell an editor is open on, which claims it. */
    | { type: "presence"; cell: CellRef | null }
    /**
     * The cell a cursor is resting on, which claims nothing. Separate from
     * `presence` so a build that predates it drops the message rather than
     * reading a parked cursor as a claim and refusing its own debater's
     * keystrokes.
     */
    | { type: "cursor"; cell: CellRef | null }
    | { type: "bye" };

/**
 * A field long enough for any round name or display name a debater types, and
 * short enough that a peer cannot use one as somewhere to put a payload.
 */
const MAX_FIELD = 256;

type Hello = Extract<WireMessage, { type: "hello" }>;
type HelloAck = Extract<WireMessage, { type: "helloAck" }>;
type DocMessage = Extract<WireMessage, { type: "state" | "delta" }>;
type VectorMessage = Extract<WireMessage, { type: "vector" }>;
type PositionMessage = Extract<WireMessage, { type: "presence" | "cursor" }>;

function isRecord(value: unknown): value is Record<string, unknown> {
    return typeof value === "object" && value !== null && !Array.isArray(value);
}

function isField(value: unknown): value is string {
    return typeof value === "string" && value.length <= MAX_FIELD;
}

function isOptionalField(value: unknown): boolean {
    return value === undefined || isField(value);
}

/** A non-negative integer, which is what a protocol major and an index both are. */
function isCount(value: unknown): value is number {
    return typeof value === "number" && Number.isInteger(value) && value >= 0;
}

function isRole(value: unknown): value is Role {
    return value === "partner" || value === "coach";
}

function isStamp(value: unknown): value is Stamp {
    return (
        isRecord(value) &&
        typeof value.ms === "number" &&
        typeof value.counter === "number" &&
        typeof value.actor === "string"
    );
}

/** A cell a peer claims, and nothing that only looks like one. */
export function isCellRef(value: unknown): value is CellRef {
    return isRecord(value) && isField(value.sheetId) && isCount(value.col) && isCount(value.row);
}

function isHello(m: Record<string, unknown>): m is Hello {
    return (
        m.type === "hello" &&
        isCount(m.protocol) &&
        isField(m.app) &&
        isField(m.endpointId) &&
        isField(m.roundId) &&
        isRole(m.role) &&
        Array.isArray(m.capabilities) &&
        m.capabilities.every(isField) &&
        isOptionalField(m.ticket) &&
        isOptionalField(m.label) &&
        isOptionalField(m.name)
    );
}

function isHelloAck(m: Record<string, unknown>): m is HelloAck {
    if (m.type !== "helloAck") return false;
    if (m.ok === true) return isOptionalField(m.name) && (m.role === undefined || isRole(m.role));
    return m.ok === false && isField(m.reason);
}

/**
 * A document's outline rather than its contents. The merge is written to
 * survive a register it does not recognize, but the vector walks `round` and
 * `sheets` the moment the message lands and throws when either is absent.
 */
function isDocMessage(m: Record<string, unknown>): m is DocMessage {
    if (m.type !== "state" && m.type !== "delta") return false;
    const doc = m.doc;
    return (
        isRecord(doc) &&
        typeof doc.roundId === "string" &&
        isRecord(doc.round) &&
        isRecord(doc.sheets)
    );
}

function isVector(m: Record<string, unknown>): m is VectorMessage {
    return m.type === "vector" && isRecord(m.seen) && Object.values(m.seen).every(isStamp);
}

function isPosition(m: Record<string, unknown>): m is PositionMessage {
    return (m.type === "presence" || m.type === "cursor") && (m.cell === null || isCellRef(m.cell));
}

/**
 * The message a peer sent, or null for anything that does not conform to its
 * variant.
 *
 * Everything above the transport dereferences these fields without asking: a
 * `state` carrying no document throws where the vector is taken, and a `hello`
 * whose ticket is an array throws inside the secret comparison. A peer chooses
 * every byte of what it sends, so the shape is established at the edge and a
 * message that is not one is dropped rather than acted on.
 */
export function parseWireMessage(raw: unknown): WireMessage | null {
    if (!isRecord(raw)) return null;
    switch (raw.type) {
        case "hello":
            return isHello(raw) ? raw : null;
        case "helloAck":
            return isHelloAck(raw) ? raw : null;
        case "state":
        case "delta":
            return isDocMessage(raw) ? raw : null;
        case "vector":
            return isVector(raw) ? raw : null;
        case "presence":
        case "cursor":
            return isPosition(raw) ? raw : null;
        case "bye":
            return { type: "bye" };
        default:
            return null;
    }
}

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
    /**
     * No secret rides along. The transport authenticates the key and nothing
     * else; a ticket is spent in the hello, above this line, and a parameter
     * here would tell a reader the dial itself was authorized.
     */
    dial(endpointId: string): Promise<PeerConn>;
    stop(): Promise<void>;
}

/** Anything that can hand back a link for a configuration. */
export type PeerLinkFactory = (config: PeerLinkConfig) => Promise<PeerLink>;

/**
 * The transport for a session: iroh, in the desktop shell, and nowhere else.
 * Resolved per call rather than cached, because a session binds an endpoint
 * and stopping it must actually release it.
 *
 * There is no web answer to this and there should not be one. Shared editing
 * is an iroh endpoint, which a browser cannot bind; a stand-in that satisfied
 * the port would mint tickets nobody can redeem and report a session to a
 * debater who has none. `collabLive()` is what keeps this from being called
 * off the desktop, and the throw is what says so if that ever slips.
 */
export async function createPeerLinkFor(config: PeerLinkConfig): Promise<PeerLink> {
    if (!isDesktop()) throw new Error("Shared editing needs the desktop app");
    // Dynamic so the web bundle never pulls in Tauri's JS API, matching how
    // every other desktop touchpoint is gated.
    const mod = await import("./peerLinkTauri");
    return mod.createPeerLink(config);
}
