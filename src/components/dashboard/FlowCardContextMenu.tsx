"use client";

import {
    ContextMenu,
    ContextMenuContent,
    ContextMenuItem,
    ContextMenuSub,
    ContextMenuSubContent,
    ContextMenuSubTrigger,
    ContextMenuTrigger,
} from "@/components/ui/context-menu";

import { exportFlowAs, trashFlow } from "./flowActions";

export interface FlowCardContextMenuProps {
    id: string;
    onViewDetails: (id: string) => void;
    onChanged: () => void;
    children: React.ReactNode;
}

/**
 * Right-click affordance for a flow card: same actions as the card's kebab
 * menu, replacing the browser's native context menu.
 */
export default function FlowCardContextMenu({
    id,
    onViewDetails,
    onChanged,
    children,
}: FlowCardContextMenuProps) {
    return (
        <ContextMenu>
            <ContextMenuTrigger data-testid={`context-trigger-${id}`}>
                {children}
            </ContextMenuTrigger>
            <ContextMenuContent>
                <ContextMenuItem
                    data-testid={`context-details-${id}`}
                    onSelect={() => onViewDetails(id)}
                >
                    View details
                </ContextMenuItem>
                <ContextMenuSub>
                    <ContextMenuSubTrigger>Export</ContextMenuSubTrigger>
                    <ContextMenuSubContent>
                        <ContextMenuItem onSelect={() => void exportFlowAs(id, "json")}>
                            JSON
                        </ContextMenuItem>
                        <ContextMenuItem onSelect={() => void exportFlowAs(id, "excel")}>
                            Excel
                        </ContextMenuItem>
                    </ContextMenuSubContent>
                </ContextMenuSub>
                <ContextMenuItem
                    data-testid={`context-delete-${id}`}
                    onSelect={() => void trashFlow(id, onChanged)}
                    className="text-destructive"
                >
                    Delete
                </ContextMenuItem>
            </ContextMenuContent>
        </ContextMenu>
    );
}
