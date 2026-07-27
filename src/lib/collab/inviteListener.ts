/**
 * Being reachable between rounds.
 *
 * A session binds an endpoint for the round it is holding, so with no round
 * open there is nothing for a partner to dial and an invite lands nowhere.
 * This binds the same endpoint for one purpose only: hearing a saved contact
 * offer a round, so the corner can say so.
 *
 * It joins nothing and answers nothing else. A dialler who is not in the
 * contact table gets no reply at all, not even a refusal, because an
 * EndpointId is permanent and every peer who has ever shared with this install
 * holds one. And the master switch gates it exactly like a session: off, this
 * returns null before the link factory is ever called.
 */

import type { Contacts } from "./contacts";
import { collabSettings, type CollabSettings } from "./enabled";
import { INVITED, inviteFrom, type InviteNotice } from "./invite";
import type { PeerLinkFactory } from "./peerLink";

export interface InviteListenerDeps {
    createLink: PeerLinkFactory;
    contacts(): Contacts;
    onInvite(notice: InviteNotice): void;
    settings?: () => CollabSettings;
}

export interface InviteListener {
    endpointId: string;
    stop(): Promise<void>;
}

/** Null when shared editing is off, which binds no endpoint at all. */
export async function startInviteListener(
    deps: InviteListenerDeps,
): Promise<InviteListener | null> {
    const settings = (deps.settings ?? collabSettings)();
    if (!settings.enabled) return null;

    const link = await deps.createLink({ discovery: "mdns", relay: settings.relay });
    const endpointId = await link.endpointId();
    let stopped = false;

    await link.listen((conn) => {
        if (stopped) return;
        let greeted = false;
        conn.onMessage((msg) => {
            if (greeted) return;
            greeted = true;
            // No round is held here, so every hello is about someone else's.
            const notice = inviteFrom(msg, deps.contacts(), null, conn.id);
            if (notice) {
                deps.onInvite(notice);
                conn.send({ type: "helloAck", ok: false, reason: INVITED });
            }
            conn.close();
        });
    });

    return {
        endpointId,
        async stop() {
            stopped = true;
            await link.stop();
        },
    };
}
