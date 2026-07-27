"use client";

import { useEffect, useRef, useState } from "react";

import { Dialog, DialogContent, DialogHeader, DialogTitle } from "@/components/ui/dialog";
import type { Contacts } from "@/lib/collab/contacts";
import type { Role } from "@/lib/collab/types";
import { useContactPicker } from "@/lib/store/useContactPicker";
import { cn } from "@/lib/utils";

/** What each role may do, in the words the chip shows a debater. */
const ROLE_LABEL: Record<Role, string> = {
    partner: "can edit",
    coach: "view only",
};

/**
 * Who to invite, asked once by `collab.invite`. Nothing is dialed until a name
 * is chosen here, and backing out with Escape dials nobody.
 */
export default function ContactPickerDialog() {
    const contacts = useContactPicker((s) => s.contacts);
    const pick = useContactPicker((s) => s.pick);
    const cancel = useContactPicker((s) => s.cancel);

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
                so each opening starts at the top without an effect to reset it. */}
            <DialogContent className="max-w-sm" data-testid="contact-picker">
                <DialogHeader>
                    <DialogTitle>Invite a partner</DialogTitle>
                </DialogHeader>
                {contacts && <Choices contacts={contacts} onPick={pick} />}
            </DialogContent>
        </Dialog>
    );
}

function Choices({ contacts, onPick }: { contacts: Contacts; onPick: (id: string) => void }) {
    const entries = Object.entries(contacts);
    const [cursor, setCursor] = useState(0);
    const rows = useRef<(HTMLButtonElement | null)[]>([]);

    // The arrow keys move focus itself rather than a drawn marker, so Enter and
    // Space activate the row the same way a Tab to it would.
    useEffect(() => {
        rows.current[cursor]?.focus();
    }, [cursor]);

    function onKeyDown(e: React.KeyboardEvent) {
        if (e.metaKey || e.ctrlKey || e.altKey) return;
        if (e.key === "ArrowDown") {
            e.preventDefault();
            setCursor((c) => (c + 1) % entries.length);
        } else if (e.key === "ArrowUp") {
            e.preventDefault();
            setCursor((c) => (c - 1 + entries.length) % entries.length);
        }
    }

    return (
        <div className="flex flex-col" onKeyDown={onKeyDown}>
            {entries.map(([endpointId, contact], i) => (
                <button
                    key={endpointId}
                    ref={(el) => {
                        rows.current[i] = el;
                    }}
                    type="button"
                    data-testid={`contact-pick-${endpointId}`}
                    onMouseEnter={() => setCursor(i)}
                    onFocus={() => setCursor(i)}
                    onClick={() => onPick(endpointId)}
                    className={cn(
                        "flex w-full items-center gap-3 rounded px-2 py-1.5 text-left text-sm outline-none",
                        i === cursor ? "bg-accent text-accent-foreground" : "",
                    )}
                >
                    <span className="min-w-0 flex-1 truncate font-medium">{contact.name}</span>
                    <span className="text-muted-foreground shrink-0 text-[11px]">
                        {ROLE_LABEL[contact.role]}
                    </span>
                </button>
            ))}
        </div>
    );
}
