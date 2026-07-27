/**
 * Who is allowed to put something on your screen.
 *
 * A saved contact's invite arrives as a corner message with a Join action, and
 * nothing happens until the receiver acts. An invite from anyone else produces
 * no UI at all: not a toast, not a chip flicker. An EndpointId is permanent
 * and every peer you have ever shared with holds yours, so an unknown dialler
 * that could raise a notification would be a way to interrupt a debater
 * mid-speech from across a tournament.
 *
 * Mutual-trust auto-join is deliberately absent. A partner who shares the
 * wrong round could otherwise pull a flow onto your screen mid-speech, and
 * nothing but your own hands changes what is in front of you.
 */

import { contactName, isKnown, type Contacts } from "./contacts";

/** Whether this dialler has earned a corner message. */
export function shouldAnnounceInvite(contacts: Contacts, endpointId: string): boolean {
    return isKnown(contacts, endpointId);
}

/** What the corner says. Naming the round is what makes it actionable. */
export function inviteToastFor(contacts: Contacts, endpointId: string, roundLabel: string): string {
    const who = contactName(contacts, endpointId);
    return roundLabel.trim() ? `${who} shared ${roundLabel}` : `${who} shared a round`;
}
