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
import { useTicketDialog } from "@/lib/store/useTicketDialog";

/**
 * The invite, on screen: shown after a share, asked for before a join.
 *
 * A copy that lands says so in a corner message and closes, because there is
 * nothing left to do with a dialog holding a ticket already on the clipboard.
 * The ticket is selectable text rather than a field, so a webview that refuses
 * the write still leaves a Cmd+C, which is what the failure path selects for.
 */
export default function TicketDialog() {
    const showing = useTicketDialog((s) => s.showing);
    const asking = useTicketDialog((s) => s.resolve) !== null;
    const submit = useTicketDialog((s) => s.submit);
    const close = useTicketDialog((s) => s.close);

    // A join waiting on a dialog that goes away with the tree would wait
    // forever, so leaving settles the request as a cancel.
    useEffect(() => close, [close]);

    return (
        <Dialog
            open={showing !== null || asking}
            onOpenChange={(open) => {
                if (!open) close();
            }}
        >
            {/* The field lives in the content, which unmounts with the dialog,
                so each opening starts empty without an effect to reset it. */}
            <DialogContent className="max-w-md" data-testid="ticket-dialog">
                {showing !== null ? (
                    <Handover ticket={showing} onDone={close} />
                ) : (
                    <Ask onSubmit={submit} />
                )}
            </DialogContent>
        </Dialog>
    );
}

function Handover({ ticket, onDone }: { ticket: string; onDone: () => void }) {
    const text = useRef<HTMLParagraphElement>(null);

    /** Puts the ticket under the caret, for the Cmd+C a refused write leaves. */
    function selectTicket() {
        const node = text.current;
        if (!node) return;
        const range = document.createRange();
        range.selectNodeContents(node);
        const selection = window.getSelection();
        selection?.removeAllRanges();
        selection?.addRange(range);
    }

    async function copy() {
        try {
            await navigator.clipboard.writeText(ticket);
            toast.success("Invite copied. It works once.");
            onDone();
        } catch {
            selectTicket();
            toast.error("Could not reach the clipboard. Press Cmd+C to copy the invite.");
        }
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
