import { useEffect } from "react";

import { useFlowStore } from "@/lib/store/useFlowStore";
import { isDesktop } from "@/lib/update/adapter";

import { syncInviteWatch } from "./runtime";

/**
 * Keeps the idle invite listener in step with the master switch for the app's
 * lifetime. The switch moving is the only thing this watches; whether a session
 * holds the endpoint is the runtime's own business.
 *
 * No-op on web, where there is no transport to bind.
 */
export function useInviteWatch(): void {
    useEffect(() => {
        if (!isDesktop()) return;
        let last = useFlowStore.getState().collabEnabled;
        void syncInviteWatch();
        return useFlowStore.subscribe(() => {
            const now = useFlowStore.getState().collabEnabled;
            if (now === last) return;
            last = now;
            void syncInviteWatch();
        });
    }, []);
}
