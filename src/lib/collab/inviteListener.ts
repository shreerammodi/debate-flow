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
 * holds one.
 *
 * Two switches gate it, not one. The master switch, exactly like a session;
 * and Listen for invites, which is its own setting because this is the only
 * route in ebb that binds an endpoint with no round in hand. Off - which is
 * the default - the app reaches the network when a debater shares or joins a
 * round and at no other moment, so a cold launch says nothing to anyone.
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

/** Null unless both the master switch and Listen for invites are on. */
export async function startInviteListener(
    deps: InviteListenerDeps,
): Promise<InviteListener | null> {
    const settings = (deps.settings ?? collabSettings)();
    if (!settings.enabled || !settings.listen) return null;

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
