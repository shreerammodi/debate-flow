"use client";

import { DotsThree } from "@phosphor-icons/react";

import {
    DropdownMenu,
    DropdownMenuContent,
    DropdownMenuItem,
    DropdownMenuSub,
    DropdownMenuSubContent,
    DropdownMenuSubTrigger,
    DropdownMenuTrigger,
} from "@/components/ui/dropdown-menu";

import { exportFlowAs, trashFlow } from "./flowActions";

export interface FlowCardMenuProps {
    id: string;
    onViewDetails: (id: string) => void;
    onChanged: () => void;
}

export default function FlowCardMenu({ id, onViewDetails, onChanged }: FlowCardMenuProps) {
    return (
        <DropdownMenu>
            <DropdownMenuTrigger asChild>
                <button
                    type="button"
                    data-testid={`kebab-${id}`}
                    aria-label="Flow actions"
                    onClick={(e) => e.stopPropagation()}
                    className="bg-accent text-muted-foreground hover:bg-accent/70 absolute top-3.5 right-3.5 z-10 flex h-7 w-7 items-center justify-center rounded-md opacity-0 transition-opacity group-focus-within:opacity-100 group-hover:opacity-100 focus-visible:opacity-100"
                >
                    <DotsThree className="size-4" />
                </button>
            </DropdownMenuTrigger>
            <DropdownMenuContent align="end" onClick={(e) => e.stopPropagation()}>
                <DropdownMenuItem
                    data-testid={`kebab-details-${id}`}
                    onSelect={() => onViewDetails(id)}
                >
                    View details
                </DropdownMenuItem>
                <DropdownMenuSub>
                    <DropdownMenuSubTrigger>Export</DropdownMenuSubTrigger>
                    <DropdownMenuSubContent>
                        <DropdownMenuItem onSelect={() => void exportFlowAs(id, "json")}>
                            JSON
                        </DropdownMenuItem>
                        <DropdownMenuItem onSelect={() => void exportFlowAs(id, "excel")}>
                            Excel
                        </DropdownMenuItem>
                    </DropdownMenuSubContent>
                </DropdownMenuSub>
                <DropdownMenuItem
                    data-testid={`kebab-delete-${id}`}
                    onSelect={() => void trashFlow(id, onChanged)}
                    className="text-destructive"
                >
                    Delete
                </DropdownMenuItem>
            </DropdownMenuContent>
        </DropdownMenu>
    );
}
