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

import type { CollabDoc } from "./types";

export const SIDECAR_VERSION = 1;

export interface Sidecar {
    version: number;
    roundId: string;
    /** Digest of the `.ebb` text this document was last in step with. */
    flowHash: string;
    peers: string[];
    doc: CollabDoc;
}

export function serializeSidecar(input: {
    roundId: string;
    flowHash: string;
    peers: string[];
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
        peers: Array.isArray(s.peers) ? s.peers.filter((p) => typeof p === "string") : [],
        doc: s.doc,
    };
}
