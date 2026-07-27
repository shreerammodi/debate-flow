/**
 * A transport that misbehaves on purpose.
 *
 * `peerLinkMemory` delivers synchronously and never fails, which is what makes
 * it a trustworthy recorder for the opt-in invariant. That same honesty makes
 * it useless for proving convergence: every message arrives, in order, at
 * once. This net holds messages in a queue instead, so a test decides when and
 * in what order they land, and can drop or partition the link in between.
 *
 * Convergence is the property being tested, so the net never silently
 * discards. A message dropped by a partition stays counted in `inFlight` until
 * the test either delivers or discards it explicitly.
 */

import type { PeerConn, PeerLink, PeerLinkConfig, WireMessage } from "@/lib/collab/peerLink";

interface Pending {
    /** The endpoint that will receive this, which is who a partition names. */
    to: string;
    from: string;
    deliver: () => void;
}

export interface FaultNet {
    create(endpointId: string): (config: PeerLinkConfig) => Promise<PeerLink>;
    /** Messages written but not yet handed to the far side. */
    inFlight(): number;
    /** Delivers everything queued, oldest first, including what queues during. */
    flush(): void;
    /**
     * Delivers queued messages in an order the seed decides. Later sends can
     * land before earlier ones, which is what a real network is allowed to do.
     */
    flushShuffled(rng: () => number): void;
    /** Nothing crosses between these two until `heal`. Sends still queue. */
    partition(a: string, b: string): void;
    heal(): void;
    /** Throws away everything queued, as a link that died mid-burst would. */
    discardInFlight(): number;
    /** Closes every live connection, as a process exit would. */
    killLinks(): void;
    reset(): void;
}

interface Endpoint {
    config: PeerLinkConfig;
    onPeer: ((peer: PeerConn) => void) | null;
}

export function createFaultNet(): FaultNet {
    const endpoints = new Map<string, Endpoint>();
    let queue: Pending[] = [];
    let cuts: { a: string; b: string }[] = [];
    let closers: (() => void)[] = [];

    function isCut(x: string, y: string): boolean {
        return cuts.some((c) => (c.a === x && c.b === y) || (c.a === y && c.b === x));
    }

    function connect(diallerId: string, listenerId: string) {
        type Side = "dialler" | "listener";
        const heard: Record<Side, ((m: WireMessage) => void)[]> = { dialler: [], listener: [] };
        const closed: Record<Side, (() => void)[]> = { dialler: [], listener: [] };
        let open = true;

        function shut(): void {
            if (!open) return;
            open = false;
            for (const cb of [...closed.dialler, ...closed.listener]) cb();
        }
        closers.push(shut);

        function side(self: Side, farId: string, selfId: string): PeerConn {
            const other: Side = self === "dialler" ? "listener" : "dialler";
            return {
                id: farId,
                connectionType: () => "direct",
                send(msg) {
                    if (!open) return;
                    // Serialized at send time, so a queued message carries the
                    // document as it was written rather than as it ends up.
                    const copy = JSON.parse(JSON.stringify(msg)) as WireMessage;
                    queue.push({
                        to: farId,
                        from: selfId,
                        deliver: () => {
                            if (!open) return;
                            for (const cb of heard[other]) cb(copy);
                        },
                    });
                },
                onMessage(cb) {
                    heard[self].push(cb);
                },
                onClose(cb) {
                    closed[self].push(cb);
                },
                close: shut,
            };
        }

        return {
            dialler: side("dialler", listenerId, diallerId),
            listener: side("listener", diallerId, listenerId),
        };
    }

    /** Everything deliverable right now, leaving partitioned messages queued. */
    function takeDeliverable(): Pending[] {
        const ready: Pending[] = [];
        const held: Pending[] = [];
        for (const p of queue) (isCut(p.from, p.to) ? held : ready).push(p);
        queue = held;
        return ready;
    }

    return {
        create(endpointId) {
            return async (config) => {
                endpoints.set(endpointId, { config, onPeer: null });
                return {
                    async endpointId() {
                        return endpointId;
                    },
                    async listen(onPeer) {
                        const self = endpoints.get(endpointId);
                        if (self) self.onPeer = onPeer;
                    },
                    async dial(target) {
                        const far = endpoints.get(target);
                        if (!far?.onPeer) throw new Error(`no peer listening at ${target}`);
                        const { dialler, listener } = connect(endpointId, target);
                        far.onPeer(listener);
                        return dialler;
                    },
                    async stop() {
                        endpoints.delete(endpointId);
                    },
                };
            };
        },

        inFlight() {
            return queue.length;
        },

        flush() {
            // A delivery can provoke a reply, and the reply is part of settling.
            for (let round = 0; round < 100; round++) {
                const ready = takeDeliverable();
                if (ready.length === 0) return;
                for (const p of ready) p.deliver();
            }
            throw new Error("flush did not settle: the peers are talking in a loop");
        },

        flushShuffled(rng) {
            for (let round = 0; round < 100; round++) {
                const ready = takeDeliverable();
                if (ready.length === 0) return;
                for (let i = ready.length - 1; i > 0; i--) {
                    const j = Math.floor(rng() * (i + 1));
                    [ready[i], ready[j]] = [ready[j], ready[i]];
                }
                for (const p of ready) p.deliver();
            }
            throw new Error("flush did not settle: the peers are talking in a loop");
        },

        partition(a, b) {
            if (!isCut(a, b)) cuts.push({ a, b });
        },

        heal() {
            cuts = [];
        },

        discardInFlight() {
            const n = queue.length;
            queue = [];
            return n;
        },

        killLinks() {
            for (const shut of closers) shut();
            closers = [];
        },

        reset() {
            endpoints.clear();
            queue = [];
            cuts = [];
            closers = [];
        },
    };
}
