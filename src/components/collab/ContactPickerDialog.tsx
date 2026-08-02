"use client";

import { useEffect, useRef, useState } from "react";

import { Dialog, DialogContent, DialogHeader, DialogTitle } from "@/components/ui/dialog";
import type { Contacts } from "@/lib/collab/contacts";
import type { Role } from "@/lib/collab/types";
import { useContactPicker, type ContactChoice } from "@/lib/store/useContactPicker";
import { cn } from "@/lib/utils";

/** The grants on offer, in the words the chip shows a debater. */
const GRANTS: { role: Role; label: string }[] = [
    { role: "editor", label: "Edit" },
    { role: "viewer", label: "View" },
];

/**
 * Who to invite and what to grant them, asked once by `collab.invite`. Nothing
 * is dialed until a grant is chosen here, and backing out with Escape dials
 * nobody.
 *
 * The grant is asked rather than remembered. A contact is a partner and carries
 * no grade, so what they may do is decided for the round in front of the
 * debater and nowhere else; a preselected answer would make the wider one the
 * thing that happens when somebody presses Enter twice.
 */
export default function ContactPickerDialog() {
    const contacts = useContactPicker((s) => s.contacts);
    const pick = useContactPicker((s) => s.pick);
    const cancel = useContactPicker((s) => s.cancel);
    const list = useRef<HTMLDivElement>(null);

    // A caller waiting on a picker that goes away with the tree would wait
    // forever, so leaving settles the request as a cancel.
    useEffect(() => cancel, [cancel]);

    return (
        <Dialog
            open={contacts !== null}
            onOpenChange={(open) => {
                if (!open) cancel();
            }}
        >
            {/* The cursor lives in the content, which unmounts with the dialog,
                so each opening starts with no cursor at all and needs no effect
                to reset one. */}
            <DialogContent className="max-w-sm" data-testid="contact-picker" initialFocus={list}>
                <DialogHeader>
                    <DialogTitle>Invite a partner</DialogTitle>
                </DialogHeader>
                {contacts && <Choices ref={list} contacts={contacts} onPick={pick} />}
            </DialogContent>
        </Dialog>
    );
}

/**
 * The saved partners, each row offering the two grants.
 *
 * The container takes focus itself so that opening the dialog has somewhere to
 * land that is not a grant: the first tabbable child is otherwise the first
 * row's Edit, and the wide grant becomes whatever an Enter commits before the
 * debater has chosen anything. The arrow handler sits on that same container
 * because a keydown aimed at the popup never reaches a child's handler, so the
 * element holding focus has to be the element listening. A Tab still walks the
 * buttons themselves, which keep their place in the tab order.
 */
function Choices({
    contacts,
    onPick,
    ref,
}: {
    contacts: Contacts;
    onPick: (choice: ContactChoice) => void;
    ref: React.Ref<HTMLDivElement>;
}) {
    const entries = Object.entries(contacts);
    /**
     * Which button the keyboard is on, and null until the debater has moved.
     *
     * A grant is not preselected: opening the picker onto a focused Edit would
     * make the wider role the thing that happens when somebody presses Enter
     * twice, and this is the gesture that decides whether a peer may write into
     * the round. Nothing holds focus, so an Enter that arrives before a choice
     * commits nothing and leaves the picker open.
     */
    const [cursor, setCursor] = useState<number | null>(null);
    const stops = useRef<(HTMLButtonElement | null)[]>([]);

    // The arrow keys move focus itself rather than a drawn marker, so Enter and
    // Space activate the button the cursor is on the same way a Tab to it
    // would. One stop per button, so a row of two grants takes two.
    useEffect(() => {
        if (cursor !== null) stops.current[cursor]?.focus();
    }, [cursor]);

    function onKeyDown(e: React.KeyboardEvent) {
        if (e.metaKey || e.ctrlKey || e.altKey) return;
        const count = entries.length * GRANTS.length;
        if (e.key === "ArrowDown") {
            e.preventDefault();
            // From nowhere, down lands on the first row's first grant, which is
            // where a list with no cursor starts rather than where it resumes.
            setCursor((c) => (c === null ? 0 : (c + GRANTS.length) % count));
        } else if (e.key === "ArrowUp") {
            e.preventDefault();
            setCursor((c) =>
                c === null ? count - GRANTS.length : (c - GRANTS.length + count) % count,
            );
        }
    }

    return (
        <div
            ref={ref}
            tabIndex={-1}
            role="group"
            aria-label="Saved partners"
            className="flex flex-col outline-none"
            onKeyDown={onKeyDown}
        >
            {entries.map(([endpointId, contact], row) => (
                <div key={endpointId} className="flex w-full items-center gap-2 px-2 py-1.5">
                    <span className="min-w-0 flex-1 truncate text-sm font-medium">
                        {contact.name}
                    </span>
                    {GRANTS.map(({ role, label }, column) => {
                        const stop = row * GRANTS.length + column;
                        return (
                            <button
                                key={role}
                                ref={(el) => {
                                    stops.current[stop] = el;
                                }}
                                type="button"
                                data-testid={`contact-pick-${role}-${endpointId}`}
                                onMouseEnter={() => setCursor(stop)}
                                onFocus={() => setCursor(stop)}
                                onClick={() => onPick({ endpointId, role })}
                                className={cn(
                                    "border-border shrink-0 rounded border px-2 py-0.5 text-[12px] outline-none",
                                    stop === cursor ? "bg-accent text-accent-foreground" : "",
                                )}
                            >
                                {label}
                            </button>
                        );
                    })}
                </div>
            ))}
        </div>
    );
}
