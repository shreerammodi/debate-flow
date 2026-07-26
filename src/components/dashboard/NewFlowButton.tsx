"use client";

import { Button } from "@/components/ui/button";
import {
    DropdownMenu,
    DropdownMenuContent,
    DropdownMenuItem,
    DropdownMenuSub,
    DropdownMenuSubContent,
    DropdownMenuSubTrigger,
    DropdownMenuTrigger,
} from "@/components/ui/dropdown-menu";
import { Kbd } from "@/components/ui/kbd";
import type { KeytipId } from "@/lib/dashboard/keytips";
import type { EventId } from "@/lib/format/events";
import type { Side } from "@/lib/model/types";

import { MENU_ATTR, useKeyTips } from "./keytips/KeyTipsProvider";
import { useCreateFlow } from "./useCreateFlow";

interface EventChoice {
    event: EventId;
    label: string;
    /** Keytip whose configured chord fires this item. */
    tip: KeytipId;
}

/** Events whose speaking order is fixed; picking one creates the round. */
const FIXED_ORDER: EventChoice[] = [
    { event: "policy", label: "Policy", tip: "new.policy" },
    { event: "ld", label: "Lincoln-Douglas", tip: "new.ld" },
];

const PF_ORDERS: { firstSide: Side; label: string; tip: KeytipId }[] = [
    { firstSide: "aff", label: "Aff speaks first", tip: "new.pfFirstAff" },
    { firstSide: "neg", label: "Neg speaks first", tip: "new.pfFirstNeg" },
];

export default function NewFlowButton() {
    const create = useCreateFlow();
    const { setMode, mode, keytips } = useKeyTips();
    const tips = mode === "new";

    // The menu stays uncontrolled (mouse still opens it standalone); opening
    // just tells the overlay to paint and route the item keys.
    return (
        <DropdownMenu onOpenChange={(open) => setMode(open ? "new" : "off")}>
            <DropdownMenuTrigger asChild>
                <Button size="sm" data-testid="new-flow">
                    + New flow
                </Button>
            </DropdownMenuTrigger>
            <DropdownMenuContent align="end">
                {FIXED_ORDER.map(({ event, label, tip }) => (
                    <DropdownMenuItem
                        key={event}
                        data-testid={`new-flow-${event}`}
                        {...{ [MENU_ATTR]: keytips[tip] }}
                        onSelect={() => create(event)}
                    >
                        {label}
                        {tips && keytips[tip] && <Kbd className="ml-auto">{keytips[tip]}</Kbd>}
                    </DropdownMenuItem>
                ))}
                <DropdownMenuSub>
                    <DropdownMenuSubTrigger
                        data-testid="new-flow-pf"
                        {...{ [MENU_ATTR]: keytips["new.pf"] }}
                    >
                        Public Forum
                        {tips && keytips["new.pf"] && (
                            <Kbd className="ml-auto">{keytips["new.pf"]}</Kbd>
                        )}
                    </DropdownMenuSubTrigger>
                    <DropdownMenuSubContent>
                        {PF_ORDERS.map(({ firstSide, label, tip }) => (
                            <DropdownMenuItem
                                key={firstSide}
                                data-testid={`new-flow-pf-${firstSide}`}
                                {...{ [MENU_ATTR]: keytips[tip] }}
                                onSelect={() => create("pf", firstSide)}
                            >
                                {label}
                                {tips && keytips[tip] && (
                                    <Kbd className="ml-auto">{keytips[tip]}</Kbd>
                                )}
                            </DropdownMenuItem>
                        ))}
                    </DropdownMenuSubContent>
                </DropdownMenuSub>
            </DropdownMenuContent>
        </DropdownMenu>
    );
}
