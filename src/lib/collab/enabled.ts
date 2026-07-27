/**
 * The one gate every shared editing route passes through.
 *
 * Shared editing is the only feature in ebb that reaches the network, so the
 * switch that turns it off is an invariant rather than a preference: with it
 * off the app binds no endpoint, dials no peer, publishes no discovery record,
 * and contacts no relay. Both halves of the feature ask here rather than
 * testing the conditions apart.
 */

import { useFlowStore } from "@/lib/store/useFlowStore";
import { isDesktop } from "@/lib/update/adapter";

export interface CollabSettings {
    enabled: boolean;
    relay: boolean;
}

/** What the desktop UI asks before offering a collaboration action. */
export function collabLive(): boolean {
    return isDesktop() && useFlowStore.getState().collabEnabled;
}

/** What the session asks. The transport it is handed decides the runtime. */
export function collabSettings(): CollabSettings {
    const state = useFlowStore.getState();
    return { enabled: state.collabEnabled, relay: state.collabRelayEnabled };
}
