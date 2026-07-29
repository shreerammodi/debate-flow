import { useEffect } from "react";

import { useFlowStore } from "@/lib/store/useFlowStore";
import { isDesktop } from "@/lib/update/adapter";

import { syncInviteWatch } from "./runtime";

/**
 * Keeps the idle invite listener in step with the two switches that decide
 * whether it may exist, for the app's lifetime. Those moving are the only
 * thing this watches; whether a session holds the endpoint is the runtime's
 * own business.
 *
 * The mount call binds nothing on its own: Listen for invites is off by
 * default, so a cold launch reaches the sync and stops there.
 *
 * No-op on web, where there is no transport to bind.
 */
export function useInviteWatch(): void {
    useEffect(() => {
        if (!isDesktop()) return;
        const wanted = () => {
            const s = useFlowStore.getState();
            return s.collabEnabled && s.collabListenEnabled;
        };
        let last = wanted();
        void syncInviteWatch();
        return useFlowStore.subscribe(() => {
            const now = wanted();
            if (now === last) return;
            last = now;
            void syncInviteWatch();
        });
    }, []);
}
