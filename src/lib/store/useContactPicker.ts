/**
 * useContactPicker - the invite picker's open state, and the promise behind it.
 *
 * The command layer must not reach for a component, and a dialog must not know
 * what a session is, so the two meet here: `chooseContact` opens the picker and
 * settles once the user picks a peer or backs out.
 */

import { create } from "zustand";

import type { Contacts } from "@/lib/collab/contacts";
import type { Role } from "@/lib/collab/types";

/** A partner, and what this round is granting them. */
export interface ContactChoice {
    endpointId: string;
    role: Role;
}

export interface ContactPickerState {
    /** What to pick from, and null whenever the picker is closed. */
    contacts: Contacts | null;
    resolve: ((choice: ContactChoice | null) => void) | null;
    /** Settles the open request and closes the picker. null is a cancel. */
    pick(choice: ContactChoice | null): void;
    cancel(): void;
}

export const useContactPicker = create<ContactPickerState>((set, get) => ({
    contacts: null,
    resolve: null,
    pick(choice) {
        const { resolve } = get();
        // Closed before the caller resumes, so a picker reopened from inside
        // the continuation is not torn down by this one.
        set({ contacts: null, resolve: null });
        resolve?.(choice);
    },
    cancel() {
        get().pick(null);
    },
}));

/**
 * Opens the picker and waits for the answer. A request still pending is
 * cancelled first, so no caller is left holding a promise that never settles.
 */
export function chooseContact(contacts: Contacts): Promise<ContactChoice | null> {
    useContactPicker.getState().cancel();
    const { promise, resolve } = Promise.withResolvers<ContactChoice | null>();
    useContactPicker.setState({ contacts, resolve });
    return promise;
}
