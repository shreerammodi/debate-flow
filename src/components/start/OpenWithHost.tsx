"use client";

import { useEffect } from "react";
import { toast } from "sonner";

import { navigateToFlow } from "@/lib/commands/flowNav";
import { errorMessage } from "@/lib/errorMessage";
import { noteOpened } from "@/lib/persistence/flowSession";
import { isDesktop } from "@/lib/update/adapter";

/**
 * Opens the flow the OS asked for: a double-clicked `.ebb`, an "Open With", or
 * the argv of a second launch that `tauri-plugin-single-instance` forwarded to
 * the running window.
 *
 * The request can land before the webview exists, so Rust buffers it and this
 * drains the buffer once on mount, then listens for anything that arrives
 * later. Renders nothing; mounted once in the root layout.
 */
export default function OpenWithHost() {
    useEffect(() => {
        if (!isDesktop()) return;
        let unlisten: (() => void) | undefined;
        let mounted = true;

        void (async () => {
            // Platform-only modules: the browser bundle must not pull them in.
            const [{ invoke }, { listen }] = await Promise.all([
                import("@tauri-apps/api/core"),
                import("@tauri-apps/api/event"),
            ]);

            const openPath = (path: string) => {
                void noteOpened(path);
                navigateToFlow(path);
            };

            const stop = await listen<string>("file:open", (event) => openPath(event.payload));
            if (!mounted) {
                stop();
                return;
            }
            unlisten = stop;

            try {
                // Drain after the listener is attached, so a path arriving in
                // between is emitted to a live listener rather than dropped.
                const pending = await invoke<string[]>("drain_pending_open");
                // Only the last one can win a single window; opening the rest
                // would just be a flicker on the way to it.
                if (pending.length) openPath(pending[pending.length - 1]);
            } catch (err) {
                toast.error(errorMessage(err, "Could not open that flow"));
            }
        })();

        return () => {
            mounted = false;
            unlisten?.();
        };
    }, []);

    return null;
}
