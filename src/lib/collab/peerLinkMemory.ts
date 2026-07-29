/**
 * The in-process transport the suite runs against.
 *
 * It records every call it is asked to make. That is what turns the opt-in
 * claim into a fact: with the master switch off the recorder stays empty, so
 * no endpoint was bound, no peer was dialled, no discovery record was
 * published, and no relay was contacted.
 *
 * It also narrows every line the way the desktop adapter does, because a
 * validator the suite never reaches is one that can be weakened without a
 * single test noticing. A message this drops is a message the shipping
 * transport would have dropped too.
 */

import {
    parseWireMessage,
    type PeerConn,
    type PeerLink,
    type PeerLinkConfig,
    type WireMessage,
} from "./peerLink";

export interface MemoryCall {
    op: "create" | "endpointId" | "listen" | "dial" | "stop";
    /** For a dial, the endpoint reached out to; otherwise the local one. */
    endpointId?: string;
    config?: PeerLinkConfig;
}

export interface MemoryNet {
    /** Every call made through this net, in order. */
    calls: MemoryCall[];
    /** A link factory for one endpoint, in the shape a session expects. */
    create(endpointId: string): (config: PeerLinkConfig) => Promise<PeerLink>;
    reset(): void;
}

interface Endpoint {
    config: PeerLinkConfig;
    onPeer: ((peer: PeerConn) => void) | null;
}

type Side = "dialler" | "listener";

/** The two ends of one connection, wired to each other. */
function connect(
    diallerId: string,
    listenerId: string,
    relayed: boolean,
): { dialler: PeerConn; listener: PeerConn } {
    const listeners: Record<Side, ((m: WireMessage) => void)[]> = { dialler: [], listener: [] };
    const closers: Record<Side, (() => void)[]> = { dialler: [], listener: [] };
    let open = true;

    function side(self: Side, id: string): PeerConn {
        const other: Side = self === "dialler" ? "listener" : "dialler";
        return {
            id,
            connectionType: () => (relayed ? "relayed" : "direct"),
            send(msg) {
                if (!open) return;
                // A real link carries bytes: it serializes, and a line that
                // does not conform to its variant never reaches a listener.
                const parsed = parseWireMessage(JSON.parse(JSON.stringify(msg)));
                if (!parsed) return;
                for (const cb of listeners[other]) cb(parsed);
            },
            onMessage(cb) {
                listeners[self].push(cb);
            },
            onClose(cb) {
                closers[self].push(cb);
            },
            close() {
                if (!open) return;
                open = false;
                for (const cb of [...closers.dialler, ...closers.listener]) cb();
            },
        };
    }

    // Each handle names the far side, which is what a peer list displays.
    return { dialler: side("dialler", listenerId), listener: side("listener", diallerId) };
}

export function createMemoryNet(): MemoryNet {
    const endpoints = new Map<string, Endpoint>();
    const calls: MemoryCall[] = [];

    return {
        calls,
        reset() {
            endpoints.clear();
            calls.length = 0;
        },
        create(endpointId) {
            return async (config) => {
                calls.push({ op: "create", endpointId, config });
                endpoints.set(endpointId, { config, onPeer: null });
                return {
                    async endpointId() {
                        calls.push({ op: "endpointId", endpointId });
                        return endpointId;
                    },
                    async listen(onPeer) {
                        calls.push({ op: "listen", endpointId });
                        const self = endpoints.get(endpointId);
                        if (self) self.onPeer = onPeer;
                    },
                    async dial(target) {
                        calls.push({ op: "dial", endpointId: target });
                        const far = endpoints.get(target);
                        if (!far?.onPeer) throw new Error(`no peer listening at ${target}`);
                        const { dialler, listener } = connect(
                            endpointId,
                            target,
                            config.relay && far.config.relay,
                        );
                        far.onPeer(listener);
                        return dialler;
                    },
                    async stop() {
                        calls.push({ op: "stop", endpointId });
                        endpoints.delete(endpointId);
                    },
                };
            };
        },
    };
}
