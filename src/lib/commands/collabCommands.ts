/**
 * The three collaboration commands, and nothing else reaches a session.
 *
 * Palette only: no chord, no menu accelerator. Flowing owns most of the letter
 * space, and a printable key bound outside `HotGrid`'s guard erases the cell
 * the debater is standing on.
 */

import { collabLive } from "@/lib/collab/enabled";
import { joinRound } from "@/lib/collab/join";
import { createPeerLinkFor } from "@/lib/collab/peerLink";
import { currentSession, endSession, startForRound } from "@/lib/collab/runtime";
import { encodeTicket } from "@/lib/collab/ticket";
import { useFlowStore } from "@/lib/store/useFlowStore";
import { getCurrentVersion } from "@/lib/update/adapter";

export interface CollabCommandDeps {
    /** Corner messages. Nothing here blocks the grid or takes focus. */
    notify(message: string): void;
    fail(message: string): void;
    /** Reads the pasted ticket. Returns null when the user backs out. */
    askForTicket(): Promise<string | null>;
    copy(text: string): Promise<void>;
    /** Routes to a flow file, for a join that landed. */
    openFlow(path: string): void;
}

/**
 * Mints a ticket for the open round and puts it on the clipboard, starting a
 * session first when none is running.
 */
export async function runShare(deps: CollabCommandDeps): Promise<void> {
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
        await deps.copy(encodeTicket(session.share("partner")));
        deps.notify("Invite copied. It works once.");
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
            deps.fail("Turn on shared editing in Settings first");
            return;
        }
        deps.notify(joined.created ? "Joined. The round is yours to keep." : "Joined.");
        deps.openFlow(joined.path);
    } catch (err) {
        deps.fail(err instanceof Error ? err.message : "Could not join that round");
    }
}

/** Drops every peer. The flow stays open and stays editable. */
export async function runEnd(deps: CollabCommandDeps): Promise<void> {
    if (!currentSession()) {
        deps.fail("No session is running");
        return;
    }
    await endSession();
    deps.notify("Session ended. The flow is still yours.");
}
