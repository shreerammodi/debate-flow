"use client";

import { useRouter, useSearchParams } from "next/navigation";
import { useEffect, useState } from "react";
import { toast } from "sonner";

import { Skeleton } from "@/components/ui/skeleton";
import { flowRouteFor } from "@/lib/commands/flowNav";
import { errorMessage } from "@/lib/errorMessage";
import { applyFlowFont } from "@/lib/fonts/applyFlowFont";
import { basename } from "@/lib/persistence/flowPaths";
import { attachFlowAutosave, noteOpened, readFlowAt } from "@/lib/persistence/flowSession";
import { useFlowStore } from "@/lib/store/useFlowStore";
import { useSaveStatus } from "@/lib/store/useSaveStatus";
import { applySideColors } from "@/lib/theme/applySideColors";

import Workspace from "./Workspace";

/**
 * AppRoot - boots the editor for the flow file named by ?path=.
 *
 * The path is the flow's identity, so it rides in the URL the way the database
 * id used to: a reload, or the relaunch after an update installs, reopens the
 * same file. Anything that cannot be read sends the user back to the start
 * screen with a reason, because a flow silently not opening is indistinguishable
 * from a flow that is gone.
 */
export default function AppRoot() {
    const router = useRouter();
    const params = useSearchParams();
    const path = params.get("path");
    const round = useFlowStore((s) => s.round);
    const flowFont = useFlowStore((s) => s.flowFont);
    const affColor = useFlowStore((s) => s.affColor);
    const negColor = useFlowStore((s) => s.negColor);
    const [loaded, setLoaded] = useState(false);

    useEffect(() => {
        applyFlowFont(flowFont);
    }, [flowFont]);

    useEffect(() => {
        applySideColors({ aff: affColor, neg: negColor });
    }, [affColor, negColor]);

    useEffect(() => {
        let mounted = true;
        const unsubscribe = attachFlowAutosave(useFlowStore, useSaveStatus.getState().report);

        const leave = () => {
            mounted = false;
            unsubscribe();
            useSaveStatus.getState().reset();
        };

        if (!path) {
            router.replace("/");
            return leave;
        }

        // Save As rewrites the URL to the file it just wrote, which the store
        // is already editing. Reloading it would discard nothing but would
        // flash the loading frame for no reason.
        if (useFlowStore.getState().docPath === path) {
            setLoaded(true);
            return leave;
        }

        readFlowAt(path)
            .then((r) => {
                if (!mounted) return;
                if (!r) {
                    toast.error(`${basename(path)} no longer exists`);
                    router.replace("/");
                    return;
                }
                const newFlow = params.get("new") != null;
                useFlowStore.getState().loadRound(r, { docPath: path, newFlow });
                void noteOpened(path);
                // Drop the one-shot marker so a later refresh loads this flow
                // as existing and restores the persisted RFD preference.
                if (newFlow) router.replace(flowRouteFor(path));
            })
            .catch((err: unknown) => {
                toast.error(errorMessage(err, "Could not open that flow"));
                router.replace("/");
            })
            .finally(() => {
                if (mounted) setLoaded(true);
            });

        return leave;
        // eslint-disable-next-line react-hooks/exhaustive-deps -- path keys the load; params is read as a one-shot snapshot
    }, [path, router]);

    if (!loaded || !round) {
        // Held frame mirroring the editor shell, so loading a round never
        // flashes a blank screen that reads as data loss.
        return (
            <div className="flex h-screen flex-col" data-testid="editor-loading">
                <div className="border-border bg-card flex h-12 flex-none items-center border-b px-4">
                    <Skeleton className="h-4 w-48" />
                </div>
                <div className="flex min-h-0 flex-1">
                    <div className="border-border bg-card w-[220px] shrink-0 space-y-2 border-r p-2">
                        <Skeleton className="h-7 w-full" />
                        <Skeleton className="h-7 w-full" />
                        <Skeleton className="h-7 w-2/3" />
                    </div>
                    <div className="flex-1 p-4">
                        <Skeleton className="h-40 w-full" />
                    </div>
                </div>
            </div>
        );
    }
    return <Workspace />;
}
