/**
 * An invitation, from the corner it arrives in to the Join that acts on it.
 *
 * Nothing here happens on its own. A notice is shown and remembered, and the
 * round only lands on this machine when the debater says so, because a partner
 * sharing the wrong round must never be able to pull a flow onto a screen
 * mid-speech.
 */

import { toast } from "sonner";

import { navigateToFlow } from "@/lib/commands/flowNav";
import { useCollabStore } from "@/lib/store/useCollabStore";
import { useFlowStore } from "@/lib/store/useFlowStore";
import { getCurrentVersion } from "@/lib/update/adapter";

import { inviteToastFor, shouldAnnounceInvite, type InviteNotice } from "./invite";
import { joinRound } from "./join";
import { createPeerLinkFor } from "./peerLink";

/** How long an invitation stays in the corner. Long enough to act on between speeches. */
const INVITE_TOAST_MS = 30_000;

export function announceInvite(notice: InviteNotice): void {
    const contacts = useFlowStore.getState().contacts;
    // The transport already refuses a stranger, and this refuses one again
    // rather than trusting a single gate with what reaches the screen.
    if (!shouldAnnounceInvite(contacts, notice.endpointId)) return;
    useCollabStore.getState().pushInvite(notice);
    toast(inviteToastFor(contacts, notice.endpointId, notice.label), {
        duration: INVITE_TOAST_MS,
        action: { label: "Join", onClick: () => void acceptInvite(notice) },
    });
}

/**
 * Takes the round the notice named: fetches it from the host, writes it as a
 * real `.ebb`, and opens it. The round's own session re-dials from there, so
 * this is the same path a pasted ticket takes minus the ticket.
 */
export async function acceptInvite(notice: InviteNotice): Promise<void> {
    try {
        const joined = await joinRound({
            invite: { endpointId: notice.endpointId, roundId: notice.roundId },
            createLink: createPeerLinkFor,
            appVersion: await getCurrentVersion(),
        });
        if (!joined) {
            toast.error("Turn on shared editing in Settings first");
            return;
        }
        useCollabStore.getState().dismissInvite(notice.endpointId, notice.roundId);
        toast.success(joined.created ? "Joined. The round is yours to keep." : "Joined.");
        navigateToFlow(joined.path);
    } catch (err) {
        toast.error(err instanceof Error ? err.message : "Could not join that round");
    }
}
