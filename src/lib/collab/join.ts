/**
 * Accepting an invitation.
 *
 * A join is deliberately short-lived: it dials the host once with the ticket,
 * takes the round's state, and puts a real `.ebb` on this machine. Then it
 * hangs up. The round opens like any other file, and its own session re-dials
 * the host, which by then knows this peer by EndpointId and needs no ticket.
 * That is one code path for joining and for every reconnect after it.
 *
 * A guest owns a real file for the same reason every peer does: a dead peer, a
 * dead network, and a dead app each cost nothing.
 */

import { emptyScouting, type FlowRound } from "@/lib/model/flow";
import { parseFlowFile, serializeFlow } from "@/lib/persistence/flowFile";
import { getFlowFs, type FlowFs } from "@/lib/persistence/flowFs";
import { suggestFilename } from "@/lib/persistence/flowPaths";
import { resolveFlowsDir } from "@/lib/persistence/flowsDir";
import { loadRecents } from "@/lib/persistence/recents";

import { projectDoc } from "./doc";
import { collabSettings, type CollabSettings } from "./enabled";
import { helloFrom } from "./handshake";
import type { PeerLinkFactory, WireMessage } from "./peerLink";
import { parseTicket } from "./ticket";
import type { CollabDoc } from "./types";

export interface JoinDeps {
    /** The pasted ticket, verbatim. */
    ticket: string;
    createLink: PeerLinkFactory;
    appVersion: string;
    settings?: () => CollabSettings;
    fs?: FlowFs;
}

export interface JoinResult {
    roundId: string;
    hostEndpointId: string;
    path: string;
    /** False when a local file already held this round. */
    created: boolean;
}

/** The local file holding this round, if there is one. */
async function findExisting(fs: FlowFs, roundId: string): Promise<string | null> {
    for (const recent of await loadRecents(fs)) {
        try {
            const snapshot = await fs.readFlow(recent.path);
            if (!snapshot) continue;
            if (parseFlowFile(snapshot.text).id === roundId) return recent.path;
        } catch {
            // A recent that no longer parses is not a match, and not a reason
            // to fail a join.
        }
    }
    return null;
}

/** Null when shared editing is off; throws with a reason the corner can show. */
export async function joinRound(deps: JoinDeps): Promise<JoinResult | null> {
    const settings = (deps.settings ?? collabSettings)();
    if (!settings.enabled) return null;

    const ticket = parseTicket(deps.ticket);
    if (!ticket) throw new Error("That does not look like an ebb ticket");

    const link = await deps.createLink({
        discovery: "mdns",
        // Both sides have to agree before a relay carries anything.
        relay: settings.relay && ticket.relay,
    });

    try {
        const endpointId = await link.endpointId();
        const conn = await link.dial(ticket.endpointId, deps.ticket);

        const doc = await new Promise<CollabDoc>((resolve, reject) => {
            conn.onMessage((msg: WireMessage) => {
                if (msg.type === "helloAck" && !msg.ok) {
                    reject(new Error(msg.reason));
                    return;
                }
                // The host opens with the whole document, which is the round.
                if (msg.type === "state") resolve(msg.doc);
            });
            conn.onClose(() => reject(new Error("The host hung up")));
            conn.send(
                helloFrom({
                    endpointId,
                    roundId: ticket.roundId,
                    role: ticket.role,
                    appVersion: deps.appVersion,
                    ticket: ticket.secret,
                }),
            );
        });

        const io = deps.fs ?? (await getFlowFs());
        const existing = await findExisting(io, doc.roundId);

        const now = Date.now();
        const base: FlowRound = {
            id: doc.roundId,
            createdAt: now,
            updatedAt: now,
            scouting: emptyScouting(),
            sheets: [],
        };
        const round = projectDoc(doc, base);

        if (existing)
            return {
                roundId: doc.roundId,
                hostEndpointId: ticket.endpointId,
                path: existing,
                created: false,
            };

        const dir = await resolveFlowsDir(io);
        const path = await io.createFlow(dir, suggestFilename(round), serializeFlow(round));
        return { roundId: doc.roundId, hostEndpointId: ticket.endpointId, path, created: true };
    } finally {
        // The round's own session owns the peer from here.
        await link.stop();
    }
}
