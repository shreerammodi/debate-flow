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
 * A join asks a second question when the invite it is redeeming names a round
 * already on this disk: admitting the issuer to a round the debater already
 * holds is theirs to grant, not something a pasted string settles. It shares
 * this surface and the one pending slot, so a share opening over either
 * question still settles it.
 *
 * The command layer must not reach for a component, and a dialog must not know
 * what a session is, so the two meet here, as they do for the contact picker.
 */

import { create } from "zustand";

/** What a rejoin confirmation names: the round on this disk, and who wants in. */
export interface RejoinAsk {
    /** The round as the debater knows it, taken from their own copy of it. */
    round: string;
    /** The peer asking. Shown under the name saved for them, when there is one. */
    endpointId: string;
}

export interface TicketDialogState {
    open: boolean;
    /** Which half is on screen. Survives a close, for the exit animation. */
    mode: "show" | "ask" | "rejoin";
    /** The invite last handed over. Survives a close for the same reason. */
    ticket: string;
    /** What the last confirmation asked about. Survives a close for the same reason. */
    rejoin: RejoinAsk | null;
    /** Set while a join is waiting for an answer. */
    resolve: ((answer: string | true | null) => void) | null;
    /** Puts a minted ticket on screen. */
    show(ticket: string): void;
    /** Settles a pending question and closes. null is a cancel; true grants a rejoin. */
    submit(answer: string | true | null): void;
    close(): void;
}

export const useTicketDialog = create<TicketDialogState>((set, get) => ({
    open: false,
    mode: "show",
    ticket: "",
    rejoin: null,
    resolve: null,
    show(ticket) {
        // A share while a join is open would strand the join's caller.
        get().submit(null);
        set({ open: true, mode: "show", ticket });
    },
    submit(answer) {
        const { resolve } = get();
        // Closed before the caller resumes, so a dialog reopened from inside
        // the continuation is not torn down by this one.
        set({ open: false, resolve: null });
        resolve?.(answer);
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
    const { promise, resolve } = Promise.withResolvers<string | true | null>();
    useTicketDialog.setState({ open: true, mode: "ask", resolve });
    // One slot serves both questions, so the grant a rejoin answers with is
    // narrowed away here rather than widening what a paste can return.
    return promise.then((answer) => (typeof answer === "string" ? answer : null));
}

/**
 * Asks whether a peer belongs in a round this install already holds, and waits
 * for the answer. Anything still pending is cancelled first, for the reason
 * `askForTicket` does it.
 */
export function askToRejoin(ask: RejoinAsk): Promise<boolean> {
    useTicketDialog.getState().close();
    const { promise, resolve } = Promise.withResolvers<string | true | null>();
    useTicketDialog.setState({ open: true, mode: "rejoin", rejoin: ask, resolve });
    // Only the grant adds the peer, so a cancel, a dismissal, and the dialog
    // leaving the tree all leave the round as it was.
    return promise.then((answer) => answer === true);
}

export function showTicket(ticket: string): void {
    useTicketDialog.getState().show(ticket);
}
