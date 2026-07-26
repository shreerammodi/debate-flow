"use client";

import { ContextMenu as ContextMenuPrimitive } from "@base-ui/react/context-menu";
import { CaretRight } from "@phosphor-icons/react";
import type * as React from "react";

import { asChildProps } from "@/components/ui/as-child";
import { cn } from "@/lib/utils";

function ContextMenu({ ...props }: React.ComponentProps<typeof ContextMenuPrimitive.Root>) {
    return <ContextMenuPrimitive.Root {...props} />;
}

function ContextMenuTrigger({
    asChild,
    children,
    ...props
}: React.ComponentProps<typeof ContextMenuPrimitive.Trigger> & { asChild?: boolean }) {
    return (
        <ContextMenuPrimitive.Trigger
            data-slot="context-menu-trigger"
            {...props}
            {...asChildProps(asChild, children)}
        />
    );
}

function ContextMenuContent({
    className,
    ...props
}: React.ComponentProps<typeof ContextMenuPrimitive.Popup>) {
    return (
        <ContextMenuPrimitive.Portal>
            <ContextMenuPrimitive.Positioner className="outline-hidden">
                <ContextMenuPrimitive.Popup
                    data-slot="context-menu-content"
                    className={cn(
                        "z-50 max-h-(--available-height) min-w-[8rem] origin-(--transform-origin) overflow-x-hidden overflow-y-auto rounded-md border bg-popover p-1 text-popover-foreground shadow-md ease-out-quart data-[closed]:duration-100 data-[closed]:animate-out data-[closed]:fade-out-0 motion-safe:data-[closed]:zoom-out-95 data-[open]:animate-in data-[open]:fade-in-0 motion-safe:data-[open]:zoom-in-95",
                        className,
                    )}
                    {...props}
                />
            </ContextMenuPrimitive.Positioner>
        </ContextMenuPrimitive.Portal>
    );
}

function ContextMenuItem({
    className,
    inset,
    variant = "default",
    onSelect,
    onClick,
    ...props
}: Omit<React.ComponentProps<typeof ContextMenuPrimitive.Item>, "onClick"> & {
    inset?: boolean;
    variant?: "default" | "destructive";
    // Base UI's Item exposes selection as onClick (fired for keyboard and
    // pointer alike); expose it here as onSelect for a stable call-site API.
    onSelect?: () => void;
    onClick?: React.ComponentProps<typeof ContextMenuPrimitive.Item>["onClick"];
}) {
    return (
        <ContextMenuPrimitive.Item
            data-slot="context-menu-item"
            data-inset={inset}
            data-variant={variant}
            onClick={onSelect ?? onClick}
            className={cn(
                "relative flex cursor-default items-center gap-2 rounded-sm px-2 py-1.5 text-sm outline-hidden select-none data-[highlighted]:bg-accent data-[highlighted]:text-accent-foreground data-[disabled]:pointer-events-none data-[disabled]:opacity-50 data-[inset]:pl-8 data-[variant=destructive]:text-destructive data-[variant=destructive]:data-[highlighted]:bg-destructive/10 data-[variant=destructive]:data-[highlighted]:text-destructive dark:data-[variant=destructive]:data-[highlighted]:bg-destructive/20 [&_svg]:pointer-events-none [&_svg]:shrink-0 [&_svg:not([class*='size-'])]:size-4 [&_svg:not([class*='text-'])]:text-muted-foreground data-[variant=destructive]:*:[svg]:text-destructive!",
                className,
            )}
            {...props}
        />
    );
}

function ContextMenuSeparator({
    className,
    ...props
}: React.ComponentProps<typeof ContextMenuPrimitive.Separator>) {
    return (
        <ContextMenuPrimitive.Separator
            data-slot="context-menu-separator"
            className={cn("-mx-1 my-1 h-px bg-border", className)}
            {...props}
        />
    );
}

function ContextMenuSub({
    ...props
}: React.ComponentProps<typeof ContextMenuPrimitive.SubmenuRoot>) {
    return <ContextMenuPrimitive.SubmenuRoot {...props} />;
}

function ContextMenuSubTrigger({
    className,
    inset,
    children,
    ...props
}: React.ComponentProps<typeof ContextMenuPrimitive.SubmenuTrigger> & {
    inset?: boolean;
}) {
    return (
        <ContextMenuPrimitive.SubmenuTrigger
            data-slot="context-menu-sub-trigger"
            data-inset={inset}
            className={cn(
                "flex cursor-default items-center gap-2 rounded-sm px-2 py-1.5 text-sm outline-hidden select-none data-[highlighted]:bg-accent data-[highlighted]:text-accent-foreground data-[inset]:pl-8 data-[popup-open]:bg-accent data-[popup-open]:text-accent-foreground [&_svg]:pointer-events-none [&_svg]:shrink-0 [&_svg:not([class*='size-'])]:size-4 [&_svg:not([class*='text-'])]:text-muted-foreground",
                className,
            )}
            {...props}
        >
            {children}
            <CaretRight className="ml-auto size-4" />
        </ContextMenuPrimitive.SubmenuTrigger>
    );
}

function ContextMenuSubContent({
    className,
    ...props
}: React.ComponentProps<typeof ContextMenuPrimitive.Popup>) {
    return (
        <ContextMenuPrimitive.Portal>
            <ContextMenuPrimitive.Positioner className="outline-hidden">
                <ContextMenuPrimitive.Popup
                    data-slot="context-menu-sub-content"
                    className={cn(
                        "z-50 min-w-[8rem] origin-(--transform-origin) overflow-hidden rounded-md border bg-popover p-1 text-popover-foreground shadow-lg ease-out-quart data-[closed]:duration-100 motion-safe:data-[side=left]:slide-in-from-right-2 motion-safe:data-[side=right]:slide-in-from-left-2 data-[closed]:animate-out data-[closed]:fade-out-0 motion-safe:data-[closed]:zoom-out-95 data-[open]:animate-in data-[open]:fade-in-0 motion-safe:data-[open]:zoom-in-95",
                        className,
                    )}
                    {...props}
                />
            </ContextMenuPrimitive.Positioner>
        </ContextMenuPrimitive.Portal>
    );
}

export {
    ContextMenu,
    ContextMenuTrigger,
    ContextMenuContent,
    ContextMenuItem,
    ContextMenuSeparator,
    ContextMenuSub,
    ContextMenuSubTrigger,
    ContextMenuSubContent,
};
