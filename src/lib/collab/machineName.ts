/**
 * What this side calls itself, and where that name comes from.
 *
 * The hostname is the one name a debater has already set and a partner would
 * recognise, so it is the default rather than a blank field or a short
 * EndpointId. It is deliberately never written into the config file: that file
 * syncs between machines, and a baked-in hostname would follow one laptop's
 * name onto another. So the setting stays empty until someone types over it,
 * and the hostname is resolved at the moment a session needs a name.
 *
 * Asked once per run, because it cannot change while the app is open.
 */

import { useFlowStore } from "@/lib/store/useFlowStore";
import { isDesktop } from "@/lib/update/adapter";

let asked: Promise<string> | null = null;

/** The hostname, or "" when the shell cannot say and on web. */
export function machineName(): Promise<string> {
    asked ??= (async () => {
        if (!isDesktop()) return "";
        // Dynamic because the settings pane that reads this is in the web
        // bundle too, where Tauri's API does not exist.
        try {
            const { invoke } = await import("@tauri-apps/api/core");
            return await invoke<string>("machine_name");
        } catch {
            return "";
        }
    })();
    return asked;
}

/** Forgets the cached answer. For tests, which drive more than one shell. */
export function clearMachineName(): void {
    asked = null;
}

/**
 * The name a session carries: what the debater typed, or the machine's own
 * name when they typed nothing. Empty when there is neither, which greets a
 * peer with no name at all rather than with a placeholder they might save.
 */
export async function broadcastName(): Promise<string> {
    const chosen = useFlowStore.getState().collabName.trim();
    return chosen || (await machineName());
}
