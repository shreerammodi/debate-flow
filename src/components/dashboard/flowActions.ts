import { toast } from "sonner";

import { runExport } from "@/lib/export/run";
import { loadFlow, restoreFlow, softDeleteFlow } from "@/lib/persistence/flowPersistence";

/** Actions shared by the card kebab menu and the card context menu. */

export async function exportFlowAs(id: string, fmt: "json" | "excel") {
    const round = await loadFlow(id);
    if (!round) return;
    try {
        await runExport(round, fmt);
    } catch (err) {
        toast.error(`Export failed: ${err instanceof Error ? err.message : "unknown error"}`);
    }
}

export async function trashFlow(id: string, onChanged: () => void) {
    await softDeleteFlow(id);
    onChanged();
    toast("Flow moved to trash", {
        action: {
            label: "Undo",
            onClick: async () => {
                await restoreFlow(id);
                onChanged();
            },
        },
    });
}
