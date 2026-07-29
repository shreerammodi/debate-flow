"use client";

import { useEffect } from "react";
import { toast } from "sonner";

import { saveOpenFlow } from "@/lib/commands/fileCommands";
import { isDesktop } from "@/lib/update/adapter";
import { listenHere } from "@/lib/windowEvents";

/**
 * Answers the shell's flush request before the process exits.
 *
 * Quitting or closing the window used to end the process immediately, taking
 * whatever autosave had not yet written with it. Rust now holds the exit and
 * emits `app:flush`; this writes the open flow and reports back. Saying the
 * write failed cancels the exit, so a full disk or an ejected drive keeps the
 * round on screen instead of destroying it on the way out.
 *
 * Renders nothing; mounted once in the root layout.
 */
export default function QuitGuard() {
    useEffect(() => {
        if (!isDesktop()) return;
        let unlisten: (() => void) | undefined;
        let mounted = true;

        void (async () => {
            // Platform-only module: the browser bundle must not pull it in.
            const { invoke } = await import("@tauri-apps/api/core");

            // Closing one window flushes only that window, so the shell aims its
            // request; a full quit broadcasts, which reaches this listener too.
            const stop = await listenHere("app:flush", () => {
                void (async () => {
                    let saved = false;
                    try {
                        saved = await saveOpenFlow();
                    } catch {
                        saved = false;
                    }
                    if (!saved) {
                        toast.error(
                            "This flow could not be saved, so ebb stayed open. Free up space, reconnect the drive, or use Save As to put it somewhere else.",
                        );
                    }
                    await invoke("finish_quit", { saved });
                })();
            });

            if (!mounted) {
                stop();
                return;
            }
            unlisten = stop;
        })();

        return () => {
            mounted = false;
            unlisten?.();
        };
    }, []);

    return null;
}
