/**
 * The replica, made durable.
 *
 * One file per round, holding the CollabDoc, the peers the round knows, and a
 * hash of the `.ebb` it belongs to. On open a matching hash recovers the
 * replica; a missing, stale, or malformed one seeds from the file instead. The
 * sidecar is therefore an optimization and can never be a source of
 * corruption, which is why every failure below returns the same null.
 *
 * Without it one hole stays open: an app restart while diverged re-derives
 * ranks from position, so two rows independently inserted at one index collide
 * and last-writer-wins eats a cell.
 */

import { isEndpointId } from "./contacts";
import type { CollabDoc } from "./types";

/**
 * Bumped whenever a field the admission rules read starts carrying meaning,
 * because an older file parses with that field absent and so reads as the
 * widest case. Version 2 is what `coaches` costs: a version 1 file holds
 * membership with no grades, and every peer it remembers would come back a
 * partner. An unknown version is discarded, so a bump costs a re-seed of the
 * replica and never a promotion.
 */
export const SIDECAR_VERSION = 2;

export interface Sidecar {
    version: number;
    roundId: string;
    /** Digest of the `.ebb` text this document was last in step with. */
    flowHash: string;
    peers: string[];
    /** Of those peers, the ones admitted read-only. A grant the contact table never saw. */
    coaches: string[];
    doc: CollabDoc;
}

export function serializeSidecar(input: {
    roundId: string;
    flowHash: string;
    peers: string[];
    coaches: string[];
    doc: CollabDoc;
}): string {
    const sidecar: Sidecar = { version: SIDECAR_VERSION, ...input };
    return JSON.stringify(sidecar);
}

function isDoc(value: unknown): value is CollabDoc {
    if (value === null || typeof value !== "object") return false;
    const doc = value as Partial<CollabDoc>;
    return (
        typeof doc.roundId === "string" &&
        typeof doc.round === "object" &&
        doc.round !== null &&
        typeof doc.sheets === "object" &&
        doc.sheets !== null
    );
}

/**
 * Whatever of a stored list is still an id iroh could parse back into a key.
 * Anything else is a hand edit or a peer's junk, and every entry here is
 * dialled on the next open.
 */
function endpointIds(value: unknown): string[] {
    if (!Array.isArray(value)) return [];
    return value.filter((p): p is string => typeof p === "string" && isEndpointId(p));
}

/** The recovered sidecar, or null for every reason it cannot be trusted. */
export function parseSidecar(
    text: string | null,
    roundId: string,
    flowHash: string,
): Sidecar | null {
    if (!text) return null;
    let raw: unknown;
    try {
        raw = JSON.parse(text);
    } catch {
        return null;
    }
    if (raw === null || typeof raw !== "object" || Array.isArray(raw)) return null;
    const s = raw as Partial<Sidecar>;
    if (s.version !== SIDECAR_VERSION) return null;
    if (s.roundId !== roundId) return null;
    if (s.flowHash !== flowHash) return null;
    if (!isDoc(s.doc)) return null;
    return {
        version: SIDECAR_VERSION,
        roundId,
        flowHash,
        peers: endpointIds(s.peers),
        coaches: endpointIds(s.coaches),
        doc: s.doc,
    };
}
