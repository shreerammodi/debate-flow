/**
 * This install's own EndpointId, without binding anything.
 *
 * The id is the public half of the key in the identity file, so the shell can
 * answer it off the disk. Settings shows it so a partner can save this machine
 * as a contact before either side has a round to share, and that must not be a
 * reason to put a socket on the network: reading it touches nothing.
 *
 * Asked once per run, because the identity cannot change while the app is open.
 */

import { isDesktop } from "@/lib/update/adapter";

let asked: Promise<string> | null = null;

/** The id, or "" when the shell cannot say and on web. */
export function myEndpointId(): Promise<string> {
    asked ??= (async () => {
        if (!isDesktop()) return "";
        // Dynamic because the settings pane that reads this is in the web
        // bundle too, where Tauri's API does not exist.
        try {
            const { invoke } = await import("@tauri-apps/api/core");
            return await invoke<string>("collab_endpoint_id");
        } catch {
            return "";
        }
    })();
    return asked;
}

/** Forgets the cached answer. For tests, which drive more than one shell. */
export function clearMyEndpointId(): void {
    asked = null;
}
