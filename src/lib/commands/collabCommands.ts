/**
 * The three collaboration commands, and nothing else reaches a session.
 *
 * Palette only: no chord, no menu accelerator. Flowing owns most of the letter
 * space, and a printable key bound outside `HotGrid`'s guard erases the cell
 * the debater is standing on.
 */

import { contactName, type Contacts } from "@/lib/collab/contacts";
import { collabLive } from "@/lib/collab/enabled";
import { joinRound } from "@/lib/collab/join";
import { createPeerLinkFor } from "@/lib/collab/peerLink";
import { currentSession, endSession, inviteContact, startForRound } from "@/lib/collab/runtime";
import { encodeTicket } from "@/lib/collab/ticket";
import type { Role } from "@/lib/collab/types";
import type { ContactChoice } from "@/lib/store/useContactPicker";
import { useFlowStore } from "@/lib/store/useFlowStore";
import { getCurrentVersion } from "@/lib/update/adapter";

export interface CollabCommandDeps {
    /**
     * Picks a saved contact and what to grant them on this round, or null when
     * the user backs out.
     */
    chooseContact?(contacts: Contacts): Promise<ContactChoice | null>;
    /** Corner messages. Nothing here blocks the grid or takes focus. */
    notify(message: string): void;
    fail(message: string): void;
    /** Reads the ticket a guest was handed. Returns null when they back out. */
    askForTicket(): Promise<string | null>;
    /** Hands the minted ticket to the user, to send however they like. */
    presentTicket(ticket: string): void;
    /** Routes to a flow file, for a join that landed. */
    openFlow(path: string): void;
}

/**
 * Mints a ticket for the open round and puts it in front of the user, starting
 * a session first when none is running. A view-only ticket grants its holder
 * the round as it unfolds and nothing more: the host drops the writes that
 * come back from it.
 */
export async function runShare(deps: CollabCommandDeps, role: Role = "editor"): Promise<void> {
    if (!collabLive()) {
        deps.fail("Turn on shared editing in Settings first");
        return;
    }
    const round = useFlowStore.getState().round;
    if (!round) {
        deps.fail("Open a flow to share it");
        return;
    }
    try {
        const session = currentSession() ?? (await startForRound(round));
        if (!session) {
            deps.fail("Could not start a session");
            return;
        }
        deps.presentTicket(encodeTicket(await session.share(role)));
    } catch (err) {
        deps.fail(err instanceof Error ? err.message : "Could not share this round");
    }
}

/** Takes a pasted ticket, fetches the round, and opens the file it landed in. */
export async function runJoin(deps: CollabCommandDeps): Promise<void> {
    if (!collabLive()) {
        deps.fail("Turn on shared editing in Settings first");
        return;
    }
    const ticket = await deps.askForTicket();
    if (!ticket) return;
    try {
        const joined = await joinRound({
            ticket,
            createLink: createPeerLinkFor,
            appVersion: await getCurrentVersion(),
        });
        if (!joined) {
            // Either the switch went off behind the paste, or the debater
            // declined to admit the issuer to a round they already hold. The
            // second has had its dialog and wants no corner message.
            if (!collabLive()) deps.fail("Turn on shared editing in Settings first");
            return;
        }
        deps.notify(joined.created ? "Joined. The round is yours to keep." : "Joined.");
        deps.openFlow(joined.path);
    } catch (err) {
        deps.fail(err instanceof Error ? err.message : "Could not join that round");
    }
}

/**
 * Drops every peer. The flow stays open and stays editable.
 *
 * Wrapped like its siblings, and for the same reason twice over: it is fired
 * as `void runEnd(...)`, so a rejection here surfaces as an unhandled promise
 * and never reaches the user, and End session is a button a debater presses
 * mid-round. The session is torn down before the transport is asked to stop,
 * so a failure past that point has already left this side ended.
 */
export async function runEnd(deps: CollabCommandDeps): Promise<void> {
    if (!currentSession()) {
        deps.fail("No session is running");
        return;
    }
    try {
        await endSession();
        deps.notify("Session ended. The flow is still yours.");
    } catch (err) {
        deps.fail(err instanceof Error ? err.message : "Could not end the session cleanly");
    }
}

/** Dials a saved contact. No ticket: their EndpointId already authorizes. */
export async function runInvite(deps: CollabCommandDeps): Promise<void> {
    if (!collabLive()) {
        deps.fail("Turn on shared editing in Settings first");
        return;
    }
    const round = useFlowStore.getState().round;
    if (!round) {
        deps.fail("Open a flow to share it");
        return;
    }
    const contacts = useFlowStore.getState().contacts;
    if (Object.keys(contacts).length === 0) {
        deps.fail("No saved partners yet. Share a round once to save one.");
        return;
    }
    const choice = await deps.chooseContact?.(contacts);
    if (!choice) return;
    try {
        await inviteContact(round, choice.endpointId, choice.role);
        deps.notify(
            choice.role === "viewer"
                ? `Invited ${contactName(contacts, choice.endpointId)} to view`
                : `Invited ${contactName(contacts, choice.endpointId)} to edit`,
        );
    } catch (err) {
        deps.fail(err instanceof Error ? err.message : "Could not reach that partner");
    }
}
