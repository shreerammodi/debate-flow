/**
 * useCollabStore - what a shared session shows on screen.
 *
 * Kept separate from `useFlowStore` on purpose: a presence update must never
 * re-render the grid. The session layer writes here; <SessionChip /> reads.
 */

import { create } from "zustand";

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
    setStatus(status: CollabStatus): void;
    setPeers(peers: CollabPeerView[]): void;
    reset(): void;
}

/** One array backs every empty peer list, so a reset changes no identity. */
const NO_PEERS: CollabPeerView[] = [];

export const useCollabStore = create<CollabUiState>((set) => ({
    status: "off",
    peers: NO_PEERS,
    setStatus: (status) => set({ status }),
    setPeers: (peers) => set({ peers }),
    reset: () => set({ status: "off", peers: NO_PEERS }),
}));
