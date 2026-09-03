"use client";

import { useState } from "react";

import { Dialog, DialogContent, DialogHeader, DialogTitle } from "@/components/ui/dialog";
import { Input } from "@/components/ui/input";
import { Kbd } from "@/components/ui/kbd";
import type { EventId } from "@/lib/format/events";
import { makeFlowRound } from "@/lib/model/flow";
import type { Side } from "@/lib/model/types";
import { EBB_EXT, stem, suggestFilename } from "@/lib/persistence/flowPaths";
import { useFlowStore } from "@/lib/store/useFlowStore";
import { cn } from "@/lib/utils";

import { useCreateFlow } from "./useCreateFlow";

interface Choice {
    key: string;
    label: string;
    event: EventId;
    firstSide?: Side;
}

/**
 * Everything a round needs before it exists. Speaking order is only a question
 * for Public Forum, where the flip decides it; Policy, LD, and Parliamentary
 * fix the first speaker, so asking would be a step with one answer. Every
 * other detail - schools, debaters, tournament - is filled in later from
 * inside the round.
 *
 * Keys are matched against the whole list, so j and k stay free for cursor
 * movement.
 */
const CHOICES: Choice[] = [
    { key: "p", label: "Policy", event: "policy" },
    { key: "l", label: "Lincoln-Douglas", event: "ld" },
    { key: "a", label: "Public Forum, aff first", event: "pf", firstSide: "aff" },
    { key: "n", label: "Public Forum, neg first", event: "pf", firstSide: "neg" },
    { key: "r", label: "Parliamentary", event: "parli" },
];

type Picked = Pick<Choice, "event" | "firstSide">;

export default function NewFlowDialog() {
    const open = useFlowStore((s) => s.newFlowOpen);
    const setOpen = useFlowStore((s) => s.setNewFlowOpen);
    const create = useCreateFlow();
    const [picked, setPicked] = useState<Picked | null>(null);

    function onOpenChange(next: boolean, details?: { reason: string; cancel: () => void }) {
        // Escape on the name step goes back to the event list. Base UI hears
        // the key itself, so cancelling its close here is the only way in.
        if (!next && picked && details?.reason === "escape-key") {
            details.cancel();
            setPicked(null);
            return;
        }
        // The step outlives the popup, which Base UI keeps mounted through its
        // close animation, so it is reset on close rather than by unmounting.
        if (!next) setPicked(null);
        setOpen(next);
    }

    return (
        <Dialog open={open} onOpenChange={onOpenChange}>
            <DialogContent className="max-w-sm" data-testid="new-flow-dialog">
                <DialogHeader>
                    <DialogTitle>{picked ? "Name this flow" : "New flow"}</DialogTitle>
                </DialogHeader>
                {picked ? (
                    <NameStep
                        picked={picked}
                        onSubmit={(name) => {
                            onOpenChange(false);
                            create(picked.event, picked.firstSide, name);
                        }}
                    />
                ) : (
                    <Choices onPick={setPicked} />
                )}
            </DialogContent>
        </Dialog>
    );
}

/**
 * The filename, offered fully selected so the first keystroke replaces it and
 * Enter alone keeps it. The extension is fixed beside the field rather than in
 * it, so nothing the debater types can lose it.
 */
function NameStep({ picked, onSubmit }: { picked: Picked; onSubmit: (name: string) => void }) {
    const [name, setName] = useState(() =>
        stem(suggestFilename(makeFlowRound({ event: picked.event, firstSide: picked.firstSide }))),
    );

    return (
        <form
            className="flex items-center gap-1 font-mono text-sm"
            onSubmit={(e) => {
                e.preventDefault();
                onSubmit(name);
            }}
        >
            <Input
                autoFocus
                aria-label="Flow name"
                data-testid="new-flow-name"
                value={name}
                onChange={(e) => setName(e.target.value)}
                onFocus={(e) => e.target.select()}
                spellCheck={false}
            />
            <span className="text-muted-foreground">{EBB_EXT}</span>
        </form>
    );
}

function Choices({ onPick }: { onPick: (picked: Picked) => void }) {
    const [cursor, setCursor] = useState(0);

    function choose(choice: Choice) {
        onPick({ event: choice.event, firstSide: choice.firstSide });
    }

    function onKeyDown(e: React.KeyboardEvent) {
        if (e.metaKey || e.ctrlKey || e.altKey) return;

        const direct = CHOICES.findIndex((c) => c.key === e.key.toLowerCase());
        if (direct !== -1) {
            e.preventDefault();
            choose(CHOICES[direct]);
            return;
        }
        if (e.key === "ArrowDown" || e.key === "j") {
            e.preventDefault();
            setCursor((c) => (c + 1) % CHOICES.length);
        } else if (e.key === "ArrowUp" || e.key === "k") {
            e.preventDefault();
            setCursor((c) => (c - 1 + CHOICES.length) % CHOICES.length);
        } else if (e.key === "Enter") {
            e.preventDefault();
            choose(CHOICES[cursor]);
        }
    }

    return (
        <div className="font-mono text-sm" onKeyDown={onKeyDown}>
            {CHOICES.map((choice, i) => (
                <button
                    key={choice.key + choice.event}
                    type="button"
                    data-testid={`new-flow-${choice.event}${choice.firstSide ? `-${choice.firstSide}` : ""}`}
                    onMouseEnter={() => setCursor(i)}
                    onFocus={() => setCursor(i)}
                    onClick={() => choose(choice)}
                    className={cn(
                        "flex w-full items-center gap-3 rounded px-2 py-1.5 text-left outline-none",
                        i === cursor ? "bg-accent text-accent-foreground" : "",
                    )}
                >
                    <Kbd>{choice.key}</Kbd>
                    <span>{choice.label}</span>
                </button>
            ))}
        </div>
    );
}
