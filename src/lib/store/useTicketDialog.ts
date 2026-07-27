/**
 * useTicketDialog - the invite a share hands over, and the one a join asks for.
 *
 * A ticket cannot travel by clipboard API. The webview only grants
 * `navigator.clipboard` inside the task a click started, and a share has to
 * bind an endpoint first, so by the time there is a ticket to write the grant
 * is gone; reading is refused outright. The ticket goes on screen instead,
 * where selecting it is a real gesture and pasting into a field is the one
 * paste every platform allows.
 *
 * What the dialog is showing outlives `open`, because the dialog stays mounted
 * while it animates out. Clearing the mode on close would swap a closing share
 * for the join field, which reads as the whole screen flickering.
 *
 * The command layer must not reach for a component, and a dialog must not know
 * what a session is, so the two meet here, as they do for the contact picker.
 */

import { create } from "zustand";

export interface TicketDialogState {
    open: boolean;
    /** Which half is on screen. Survives a close, for the exit animation. */
    mode: "show" | "ask";
    /** The invite last handed over. Survives a close for the same reason. */
    ticket: string;
    /** Set while a join is waiting for a ticket to be pasted in. */
    resolve: ((ticket: string | null) => void) | null;
    /** Puts a minted ticket on screen. */
    show(ticket: string): void;
    /** Settles a pending ask and closes. null is a cancel. */
    submit(ticket: string | null): void;
    close(): void;
}

export const useTicketDialog = create<TicketDialogState>((set, get) => ({
    open: false,
    mode: "show",
    ticket: "",
    resolve: null,
    show(ticket) {
        // A share while a join is open would strand the join's caller.
        get().submit(null);
        set({ open: true, mode: "show", ticket });
    },
    submit(ticket) {
        const { resolve } = get();
        // Closed before the caller resumes, so a dialog reopened from inside
        // the continuation is not torn down by this one.
        set({ open: false, resolve: null });
        resolve?.(ticket);
    },
    close() {
        get().submit(null);
    },
}));

/**
 * Asks for a ticket and waits for the answer. A request still pending is
 * cancelled first, so no caller is left holding a promise that never settles.
 */
export function askForTicket(): Promise<string | null> {
    useTicketDialog.getState().close();
    const { promise, resolve } = Promise.withResolvers<string | null>();
    useTicketDialog.setState({ open: true, mode: "ask", resolve });
    return promise;
}

export function showTicket(ticket: string): void {
    useTicketDialog.getState().show(ticket);
}
