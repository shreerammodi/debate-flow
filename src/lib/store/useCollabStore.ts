/**
 * useCollabStore - what a shared session shows on screen.
 *
 * Kept separate from `useFlowStore` on purpose: a presence update must never
 * re-render the grid. The session layer writes here; <SessionChip /> reads.
 */

import { create } from "zustand";

import type { InviteNotice } from "@/lib/collab/invite";
import type { ShadowEntry } from "@/lib/collab/shadow";

export type CollabStatus = "off" | "connecting" | "connected" | "reconnecting";

export interface CollabPeerView {
    endpointId: string;
    /** Display name when a contact is known, else a short form of the id. */
    name: string;
    role: "partner" | "coach";
    connectionType: "direct" | "relayed";
}

export interface CollabUiState {
    status: CollabStatus;
    peers: CollabPeerView[];
    /**
     * This install's own EndpointId, once an endpoint has bound one. Stable
     * for the life of the identity file, so it is learned and never cleared:
     * a partner can be handed it whether or not anything is bound right now.
     */
    endpointId: string | null;
    /** Rounds saved contacts have offered and nobody has acted on yet. */
    invites: readonly InviteNotice[];
    /** What shadow mode has observed this session, oldest first. */
    shadowLog: readonly ShadowEntry[];
    setStatus(status: CollabStatus): void;
    setPeers(peers: CollabPeerView[]): void;
    setEndpointId(endpointId: string): void;
    pushInvite(invite: InviteNotice): void;
    dismissInvite(endpointId: string, roundId: string): void;
    pushShadow(entry: ShadowEntry): void;
    clearShadow(): void;
    reset(): void;
}

/** One array backs every empty peer list, so a reset changes no identity. */
const NO_PEERS: CollabPeerView[] = [];

const NO_INVITES: readonly InviteNotice[] = [];

const NO_SHADOW: readonly ShadowEntry[] = [];

/**
 * Nothing reads the log but a human, so a round long enough to overflow it
 * loses its oldest observations rather than growing the log without bound.
 */
const SHADOW_CAP = 200;

export const useCollabStore = create<CollabUiState>((set) => ({
    status: "off",
    peers: NO_PEERS,
    endpointId: null,
    shadowLog: NO_SHADOW,
    invites: NO_INVITES,
    setStatus: (status) => set({ status }),
    setPeers: (peers) => set({ peers }),
    setEndpointId: (endpointId) => set({ endpointId }),
    // A partner who dials twice about one round is one invitation, not two.
    pushInvite: (invite) =>
        set((s) =>
            s.invites.some(
                (i) => i.endpointId === invite.endpointId && i.roundId === invite.roundId,
            )
                ? s
                : { invites: [...s.invites, invite] },
        ),
    dismissInvite: (endpointId, roundId) =>
        set((s) => ({
            invites: s.invites.filter((i) => i.endpointId !== endpointId || i.roundId !== roundId),
        })),
    pushShadow: (entry) =>
        set((s) => ({
            shadowLog:
                s.shadowLog.length < SHADOW_CAP
                    ? [...s.shadowLog, entry]
                    : [...s.shadowLog.slice(s.shadowLog.length - SHADOW_CAP + 1), entry],
        })),
    clearShadow: () => set({ shadowLog: NO_SHADOW }),
    // The log outlives the session it recorded: ending a round is exactly when
    // someone sits down to read it.
    reset: () => set({ status: "off", peers: NO_PEERS }),
}));
