"use client";

import { useEffect, useRef, useState } from "react";
import { toast } from "sonner";

import { Button } from "@/components/ui/button";
import {
    Dialog,
    DialogContent,
    DialogDescription,
    DialogHeader,
    DialogTitle,
} from "@/components/ui/dialog";
import { copyText, selectNode } from "@/lib/clipboard";
import { useTicketDialog } from "@/lib/store/useTicketDialog";

/**
 * The invite, on screen: shown after a share, asked for before a join.
 *
 * A copy that lands says so in a corner message and closes, because there is
 * nothing left to do with a dialog holding a ticket already on the clipboard.
 * The ticket is selectable text rather than a field, so a webview that refuses
 * the write still leaves a Cmd+C, which is what the failure path selects for.
 *
 * Which half is on screen comes from the store's latched mode, not from what
 * is pending: a closing share must keep showing the ticket for the length of
 * its exit rather than swap to the join field on the way out.
 */
export default function TicketDialog() {
    const open = useTicketDialog((s) => s.open);
    const mode = useTicketDialog((s) => s.mode);
    const ticket = useTicketDialog((s) => s.ticket);
    const submit = useTicketDialog((s) => s.submit);
    const close = useTicketDialog((s) => s.close);

    // A join waiting on a dialog that goes away with the tree would wait
    // forever, so leaving settles the request as a cancel.
    useEffect(() => close, [close]);

    return (
        <Dialog
            open={open}
            onOpenChange={(next) => {
                if (!next) close();
            }}
        >
            {/* The field lives in the content, which unmounts with the dialog,
                so each opening starts empty without an effect to reset it. */}
            <DialogContent className="max-w-md" data-testid="ticket-dialog">
                {mode === "ask" ? (
                    <Ask onSubmit={submit} />
                ) : (
                    <Handover ticket={ticket} onDone={close} />
                )}
            </DialogContent>
        </Dialog>
    );
}

function Handover({ ticket, onDone }: { ticket: string; onDone: () => void }) {
    const text = useRef<HTMLParagraphElement>(null);

    async function copy() {
        if (await copyText(ticket)) {
            toast.success("Invite copied. It works once.");
            onDone();
            return;
        }
        selectNode(text.current);
        toast.error("Could not reach the clipboard. Press Cmd+C to copy the invite.");
    }

    return (
        <>
            <DialogHeader>
                <DialogTitle>Your invite</DialogTitle>
                <DialogDescription>
                    Send this to one partner. It works once, and it opens this round and nothing
                    else.
                </DialogDescription>
            </DialogHeader>
            <p
                ref={text}
                data-testid="ticket-text"
                className="border-border bg-muted/40 text-foreground max-h-32 overflow-y-auto rounded-md border p-2 font-mono text-[12px] break-all select-all"
            >
                {ticket}
            </p>
            <div className="flex justify-end">
                <Button type="button" size="sm" onClick={copy} data-testid="ticket-copy">
                    Copy
                </Button>
            </div>
        </>
    );
}

function Ask({ onSubmit }: { onSubmit: (ticket: string | null) => void }) {
    const [value, setValue] = useState("");

    return (
        <>
            <DialogHeader>
                <DialogTitle>Join a round</DialogTitle>
                <DialogDescription>Paste the invite your partner sent you.</DialogDescription>
            </DialogHeader>
            <textarea
                autoFocus
                rows={3}
                value={value}
                onChange={(e) => setValue(e.target.value)}
                onKeyDown={(e) => {
                    // Enter joins; the field is one value, not a paragraph.
                    if (e.key === "Enter" && !e.shiftKey) {
                        e.preventDefault();
                        if (value.trim()) onSubmit(value.trim());
                    }
                }}
                data-testid="ticket-input"
                aria-label="Invite"
                placeholder="ebb1:..."
                className="border-border bg-background text-foreground w-full resize-none rounded-md border p-2 font-mono text-[12px] break-all"
            />
            <div className="flex justify-end">
                <Button
                    type="button"
                    size="sm"
                    disabled={!value.trim()}
                    onClick={() => onSubmit(value.trim())}
                    data-testid="ticket-submit"
                >
                    Join
                </Button>
            </div>
        </>
    );
}
