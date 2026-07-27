/**
 * Starting and stopping a shared session.
 *
 * The gate is here and nowhere else: with the master switch off this function
 * returns null and never touches the link factory, so no endpoint is bound, no
 * peer is dialled, no discovery record is published, and no relay is
 * contacted. Later phases grow the body; they must not grow a second entry
 * point around it.
 */

import { collabSettings, type CollabSettings } from "./enabled";
import type { PeerConn, PeerLink, PeerLinkFactory } from "./peerLink";

export interface CollabSessionDeps {
    createLink: PeerLinkFactory;
    /** Injectable so a caller can drive the switch without a store write. */
    settings?: () => CollabSettings;
    /** Peers this round already knows, re-dialled silently when it opens. */
    peers?: string[];
    onPeer?: (conn: PeerConn) => void;
}

export interface CollabSession {
    endpointId: string;
    peers: PeerConn[];
    stop(): Promise<void>;
}

export async function startCollabSession(deps: CollabSessionDeps): Promise<CollabSession | null> {
    const settings = (deps.settings ?? collabSettings)();
    if (!settings.enabled) return null;

    const link: PeerLink = await deps.createLink({
        // mDNS reaches the machine across the room with no internet at all.
        // DNS discovery would publish this install to a public registry for a
        // session that is always invited by hand, so it is never an option.
        discovery: "mdns",
        relay: settings.relay,
    });
    const peers: PeerConn[] = [];

    function track(conn: PeerConn): void {
        peers.push(conn);
        conn.onClose(() => {
            const i = peers.indexOf(conn);
            if (i !== -1) peers.splice(i, 1);
        });
    }

    await link.listen((conn) => {
        track(conn);
        deps.onPeer?.(conn);
    });

    for (const target of deps.peers ?? []) {
        try {
            track(await link.dial(target));
        } catch {
            // A peer that is not up yet is ordinary. Reconnect is the session's
            // own job, and it never blocks the round.
        }
    }

    return {
        endpointId: await link.endpointId(),
        peers,
        async stop() {
            for (const conn of peers.slice()) conn.close();
            await link.stop();
        },
    };
}
