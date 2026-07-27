"use client";

import { useEffect, useRef, useState } from "react";

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
 * The text is always selectable, so handing a ticket over never depends on the
 * webview granting clipboard access. Copy tries the clipboard from inside the
 * click, which is the only context that is ever granted, and falls back to
 * selecting the ticket for a manual copy when even that is refused.
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
                {showing !== null ? <Handover ticket={showing} /> : <Ask onSubmit={submit} />}
            </DialogContent>
        </Dialog>
    );
}

function Handover({ ticket }: { ticket: string }) {
    const field = useRef<HTMLTextAreaElement>(null);
    const [copied, setCopied] = useState<"idle" | "done" | "manual">("idle");

    // Selected on open, so the ticket is one Cmd+C away even if the button is
    // never pressed and the clipboard API is never granted.
    useEffect(() => {
        field.current?.select();
    }, []);

    async function copy() {
        field.current?.select();
        try {
            await navigator.clipboard.writeText(ticket);
            setCopied("done");
        } catch {
            setCopied("manual");
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
            <textarea
                ref={field}
                readOnly
                rows={3}
                value={ticket}
                data-testid="ticket-text"
                aria-label="Invite"
                onFocus={(e) => e.currentTarget.select()}
                className="border-border bg-muted/40 text-foreground w-full resize-none rounded-md border p-2 font-mono text-[12px] break-all"
            />
            <div className="flex items-center justify-between gap-3">
                <span className="text-muted-foreground text-[12px]" data-testid="ticket-copy-hint">
                    {copied === "done"
                        ? "Copied."
                        : copied === "manual"
                          ? "Selected. Press Cmd+C to copy."
                          : ""}
                </span>
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
