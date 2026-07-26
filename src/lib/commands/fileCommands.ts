/**
 * Document-level commands: open, save, save as, reveal, close.
 *
 * These are the only commands that are asynchronous and that can fail in a way
 * the user needs told about, so unlike the editor commands they surface a toast
 * rather than failing silently. Each still no-ops when there is no open flow,
 * so the keyboard layer and the native menu can fire them unconditionally.
 */

import { toast } from "sonner";

import { getFlowFs } from "@/lib/persistence/flowFs";
import { pickFlowToOpen, saveFlowAs, saveFlowNow } from "@/lib/persistence/flowSession";
import { useFlowStore } from "@/lib/store/useFlowStore";
import { useSaveStatus } from "@/lib/store/useSaveStatus";

import { navigateToFlow, navigateToStart } from "./flowNav";

function report(fallback: string, err: unknown): void {
    toast.error(err instanceof Error ? err.message : fallback);
}

export async function openFlowFromPicker(): Promise<void> {
    try {
        const path = await pickFlowToOpen();
        if (path) navigateToFlow(path);
    } catch (err) {
        report("Could not open that flow", err);
    }
}

/**
 * Write the open flow now rather than waiting out the autosave debounce.
 * Autosave means it is already saved a half-second after every keystroke; this
 * exists because Cmd+S is muscle memory and pressing it should say so.
 */
export async function saveOpenFlow(): Promise<void> {
    const { round, docPath } = useFlowStore.getState();
    if (!round || !docPath) return;
    await saveFlowNow(docPath, round, useSaveStatus.getState().report);
}

export async function saveOpenFlowAs(): Promise<void> {
    const { round, setDocPath } = useFlowStore.getState();
    if (!round) return;
    try {
        const path = await saveFlowAs(round);
        if (!path) return;
        // Keep editing the new file, and keep the URL in step so a reload
        // reopens the copy the user is now looking at rather than the original.
        setDocPath(path);
        navigateToFlow(path);
        toast.success(`Saved to ${path}`);
    } catch (err) {
        report("Could not save that flow", err);
    }
}

export async function revealOpenFlow(): Promise<void> {
    const { docPath } = useFlowStore.getState();
    if (!docPath) return;
    try {
        const fs = await getFlowFs();
        await fs.reveal(docPath);
    } catch (err) {
        report("Could not show that flow", err);
    }
}

/** Flush any pending write before leaving, so the last edit is on disk. */
export async function closeOpenFlow(): Promise<void> {
    await saveOpenFlow();
    useFlowStore.getState().closeRound();
    navigateToStart();
}
