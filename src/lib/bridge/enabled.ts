/**
 * The one gate every CardMirror feature passes through.
 *
 * The bridge only exists inside the desktop shell, and the user can switch the
 * whole integration off in Settings then Editor. Both halves of the bridge ask
 * here rather than testing the two conditions apart; the settings and
 * cheatsheet UI mirror it with a subscription so they re-render on the toggle.
 */

import { useFlowStore } from "@/lib/store/useFlowStore";
import { isDesktop } from "@/lib/update/adapter";

export function cardmirrorLive(): boolean {
    return isDesktop() && useFlowStore.getState().cardmirrorEnabled;
}
