/**
 * Keeping one peer in step.
 *
 * Push, not poll: every committed change ships immediately, coalesced with a
 * short debounce so a typing burst becomes a few frames rather than a message
 * per keystroke. Nothing normal waits on a tick.
 *
 * The periodic vector is repair and never the primary path. It states the
 * highest stamp seen per actor; the far side replies with everything above
 * that, which is exact, needs no hashing, and recovers whatever a dropped link
 * lost.
 */

import { deltaSince, emptyVector, isEmptyDelta, vectorOf, type Vector } from "./delta";
import type { DroppedCell } from "./merge";
import type { PeerConn, WireMessage } from "./peerLink";
import { incomingDoc, outgoingDoc } from "./rfdSync";
import { compareStamps } from "./stamp";
import type { CollabDoc } from "./types";

/** A typing burst becomes a few frames at this width. */
const COALESCE_MS = 30;
/** Repair cadence. Long enough that it never competes with the push path. */
const REPAIR_MS = 5_000;

export interface SyncDeps {
    conn: PeerConn;
    /** The live replica, read at send time so a burst sends one current delta. */
    doc(): CollabDoc;
    apply(incoming: CollabDoc): DroppedCell[];
    /**
     * A coach reads only. In a star topology all traffic passes through the
     * host, so refusing inbound writes here is sufficient enforcement.
     */
    readOnly?: boolean;
    /** This peer's own EndpointId, which its notes travel under. */
    endpointId: string;
    now?: () => number;
    schedule?: (fn: () => void, ms: number) => () => void;
}

export interface PeerSync {
    onMessage(msg: WireMessage): void;
    notifyLocalChange(): void;
    /** The whole document, for a peer joining with no file of its own. */
    sendState(): void;
    stop(): void;
}

function defaultSchedule(fn: () => void, ms: number): () => void {
    const id = setTimeout(fn, ms);
    return () => clearTimeout(id);
}

export function attachSync(deps: SyncDeps): PeerSync {
    const schedule = deps.schedule ?? defaultSchedule;
    /**
     * What this peer has been sent. It starts at the origin because every
     * value seeded from a file carries that one stamp: two peers who opened
     * the same round already agree, and a peer that opened nothing is caught
     * up by an explicit state rather than by a delta.
     */
    let sent: Vector = emptyVector();
    let cancelPush: (() => void) | null = null;
    let cancelRepair: (() => void) | null = null;
    let stopped = false;

    function raiseSent(shipped: CollabDoc): void {
        for (const [actor, stamp] of Object.entries(vectorOf(shipped))) {
            const held = sent[actor];
            if (!held || compareStamps(stamp, held) > 0) sent[actor] = stamp;
        }
    }

    /** Sends everything the far side is missing, by its own account. */
    function sendDelta(seen: Vector): void {
        if (stopped) return;
        const delta = deltaSince(deps.doc(), seen);
        if (isEmptyDelta(delta)) return;
        deps.conn.send({ type: "delta", doc: outgoingDoc(delta, deps.endpointId) });
        raiseSent(delta);
    }

    function armRepair(): void {
        cancelRepair = schedule(() => {
            if (stopped) return;
            deps.conn.send({ type: "vector", seen: vectorOf(deps.doc()) });
            armRepair();
        }, REPAIR_MS);
    }
    armRepair();

    const sync: PeerSync = {
        notifyLocalChange() {
            if (stopped || cancelPush) return;
            cancelPush = schedule(() => {
                cancelPush = null;
                sendDelta(sent);
            }, COALESCE_MS);
        },

        sendState() {
            if (stopped) return;
            const doc = deps.doc();
            deps.conn.send({ type: "state", doc: outgoingDoc(doc, deps.endpointId) });
            raiseSent(doc);
        },

        onMessage(msg) {
            if (stopped) return;
            switch (msg.type) {
                case "delta":
                    if (deps.readOnly) return;
                    deps.apply(incomingDoc(msg.doc, deps.endpointId));
                    return;
                case "state": {
                    if (deps.readOnly) return;
                    // What the sender holds is what it has seen, so the reply
                    // is everything above that.
                    const theirs = vectorOf(msg.doc);
                    deps.apply(incomingDoc(msg.doc, deps.endpointId));
                    sendDelta(theirs);
                    return;
                }
                case "vector":
                    sendDelta(msg.seen);
                    return;
                default:
                    return;
            }
        },

        stop() {
            stopped = true;
            cancelPush?.();
            cancelPush = null;
            cancelRepair?.();
            cancelRepair = null;
        },
    };

    deps.conn.onMessage(sync.onMessage);
    return sync;
}
