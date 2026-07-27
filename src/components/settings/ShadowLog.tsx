"use client";

import { Warning } from "@phosphor-icons/react";
import { useMemo } from "react";

import { Button } from "@/components/ui/button";
import { type Contacts, contactName } from "@/lib/collab/contacts";
import type { DroppedCell } from "@/lib/collab/merge";
import { columnsForFlowSheet, type SpeechCol } from "@/lib/grid/flowColumns";
import { useCollabStore } from "@/lib/store/useCollabStore";
import { useFlowStore } from "@/lib/store/useFlowStore";
import { cn } from "@/lib/utils";

import SettingRow from "./SettingRow";

interface SheetLabel {
    title: string;
    cols: SpeechCol[];
}

/**
 * Every observation is logged, so a quiet link and a correct one never read
 * alike. A buried cell is the one loss a debater cannot see, so it outranks a
 * plain change.
 */
type EntryState = "quiet" | "changed" | "buried";

const ENTRY_BORDER: Record<EntryState, string> = {
    quiet: "border-border/40",
    changed: "border-border/60",
    buried: "border-warn/60",
};

/** A blank cell reads as blank rather than as nothing at all. */
function cellText(text: string): string {
    return text === "" ? "(blank)" : text;
}

/**
 * Where a cell sits, in the words the grid uses. An observation outlives the
 * round that produced it, so an unresolvable sheet falls back to its id.
 */
function sheetAndColumn(sheets: Map<string, SheetLabel>, sheetId: string, col: number): string {
    const sheet = sheets.get(sheetId);
    const name = sheet?.cols[col]?.name ?? `column ${col + 1}`;
    return `${sheet ? sheet.title : sheetId} - ${name}`;
}

function droppedLine(sheets: Map<string, SheetLabel>, contacts: Contacts, c: DroppedCell): string {
    const who = `written by ${contactName(contacts, c.writtenBy)}, deleted by ${contactName(contacts, c.deletedBy)}`;
    return `Dropped from ${sheetAndColumn(sheets, c.sheetId, c.col)}: ${cellText(c.text)} (${who})`;
}

/**
 * What shadow mode saw, for a human to read after the round. Absent until
 * shadow mode is switched on, and stays reachable once it is switched back off
 * so a recorded observation is never stranded behind the toggle.
 */
export default function ShadowLog() {
    const shadowMode = useFlowStore((s) => s.shadowMode);
    const round = useFlowStore((s) => s.round);
    const contacts = useFlowStore((s) => s.contacts);
    const shadowLog = useCollabStore((s) => s.shadowLog);
    const clearShadow = useCollabStore((s) => s.clearShadow);

    const sheets = useMemo(() => {
        const byId = new Map<string, SheetLabel>();
        if (round) {
            for (const sheet of round.sheets) {
                byId.set(sheet.id, { title: sheet.title, cols: columnsForFlowSheet(round, sheet) });
            }
        }
        return byId;
    }, [round]);

    const newestFirst = useMemo(() => shadowLog.slice().reverse(), [shadowLog]);

    if (!shadowMode && newestFirst.length === 0) return null;

    return (
        <div data-testid="shadow-log">
            <SettingRow
                title="Recorded changes"
                description="What a partner sent while shadow mode was on. Your flow is untouched."
                control={
                    newestFirst.length > 0 && (
                        <Button
                            variant="outline"
                            size="xs"
                            onClick={clearShadow}
                            data-testid="shadow-log-clear"
                        >
                            Clear
                        </Button>
                    )
                }
            >
                {newestFirst.length === 0 ? (
                    <p className="text-muted-foreground text-[12px]" data-testid="shadow-log-empty">
                        Nothing recorded yet.
                    </p>
                ) : (
                    <ul className="m-0 flex list-none flex-col gap-1.5 p-0">
                        {newestFirst.map((entry, i) => {
                            const state: EntryState =
                                entry.dropped.length > 0
                                    ? "buried"
                                    : entry.diffs.length > 0
                                      ? "changed"
                                      : "quiet";
                            return (
                                <li
                                    key={`${entry.at}-${i}`}
                                    className={cn(
                                        "flex flex-col gap-0.5 rounded-md border px-2 py-1.5",
                                        ENTRY_BORDER[state],
                                    )}
                                    data-testid="shadow-log-entry"
                                    data-state={state}
                                >
                                    <div className="text-muted-foreground flex items-center gap-2 text-[11px]">
                                        <span>{new Date(entry.at).toLocaleTimeString()}</span>
                                        <span>{contactName(contacts, entry.from)}</span>
                                    </div>
                                    {state === "quiet" && (
                                        <div className="text-muted-foreground text-[12px] leading-snug">
                                            No change to your flow.
                                        </div>
                                    )}
                                    {entry.diffs.map((d, j) => (
                                        <div
                                            key={`${d.sheetId}-${d.col}-${d.row}-${j}`}
                                            className="text-[12px] leading-snug"
                                            data-testid="shadow-log-diff"
                                        >
                                            <span className="text-muted-foreground">
                                                {`${sheetAndColumn(sheets, d.sheetId, d.col)} row ${d.row + 1}`}
                                            </span>
                                            <span className="text-foreground">
                                                {` ${cellText(d.mine)} -> ${cellText(d.theirs)}`}
                                            </span>
                                        </div>
                                    ))}
                                    {entry.dropped.map((c, j) => (
                                        <div
                                            key={`${c.sheetId}-${c.rank}-${j}`}
                                            className="text-warn flex items-start gap-1.5 text-[12px] leading-snug font-medium"
                                            data-testid="shadow-log-dropped"
                                        >
                                            <Warning
                                                size={13}
                                                aria-hidden="true"
                                                className="mt-0.5 shrink-0"
                                            />
                                            {droppedLine(sheets, contacts, c)}
                                        </div>
                                    ))}
                                </li>
                            );
                        })}
                    </ul>
                )}
            </SettingRow>
        </div>
    );
}
