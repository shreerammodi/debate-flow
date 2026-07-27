/**
 * useCollabStore - what a shared session shows on screen.
 *
 * Kept separate from `useFlowStore` on purpose: a presence update must never
 * re-render the grid. The session layer writes here; <SessionChip /> reads.
 */

import { create } from "zustand";

import type { InviteNotice } from "@/lib/collab/invite";

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
    setStatus(status: CollabStatus): void;
    setPeers(peers: CollabPeerView[]): void;
    setEndpointId(endpointId: string): void;
    pushInvite(invite: InviteNotice): void;
    dismissInvite(endpointId: string, roundId: string): void;
    reset(): void;
}

/** One array backs every empty peer list, so a reset changes no identity. */
const NO_PEERS: CollabPeerView[] = [];

const NO_INVITES: readonly InviteNotice[] = [];

export const useCollabStore = create<CollabUiState>((set) => ({
    status: "off",
    peers: NO_PEERS,
    endpointId: null,
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
    reset: () => set({ status: "off", peers: NO_PEERS }),
}));
