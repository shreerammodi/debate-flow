"use client";

import { useRouter } from "next/navigation";
import { useCallback, useEffect, useState } from "react";

import { Wordmark } from "@/components/brand/Logo";
import { Kbd } from "@/components/ui/kbd";
import { Skeleton } from "@/components/ui/skeleton";
import { openFlowFromPicker } from "@/lib/commands/fileCommands";
import { flowRouteFor } from "@/lib/commands/flowNav";
import { openExternal } from "@/lib/openExternal";
import { relativeTime } from "@/lib/start/format";
import { useFlowStore } from "@/lib/store/useFlowStore";
import { getCurrentVersion, isDesktop } from "@/lib/update/adapter";
import { cn } from "@/lib/utils";

import MigrationDialog from "./MigrationDialog";
import { useRecentFlows, type RecentEntry } from "./useRecentFlows";

const LINKS = [
    { label: "Documentation", href: "https://ebb.smodi.net/docs" },
    { label: "GitHub", href: "https://github.com/shreerammodi/ebb" },
    { label: "Shreeram Modi", href: "https://smodi.net", prefix: "Developed by " },
];

function Rule() {
    return <div className="bg-border/60 my-5 h-px w-full" />;
}

/**
 * The start screen, modelled on nvim's.
 *
 * There is no list of flows to manage here, because the filesystem is the
 * library now: three commands, the flows you were last in, and where to read
 * more. Every target is one keypress, which is the point - the screen exists to
 * be left quickly.
 */
export default function StartScreen() {
    const router = useRouter();
    const setNewFlowOpen = useFlowStore((s) => s.setNewFlowOpen);
    const setSettingsOpen = useFlowStore((s) => s.setSettingsOpen);
    const setCheatsheetOpen = useFlowStore((s) => s.setCheatsheetOpen);
    const { entries, refresh } = useRecentFlows();
    const [cursor, setCursor] = useState(0);
    const [version, setVersion] = useState(process.env.NEXT_PUBLIC_EBB_VERSION ?? "");

    useEffect(() => {
        // The packaged version is the truth on desktop; the injected constant
        // only covers the browser, where there is no Tauri runtime to ask.
        if (!isDesktop()) return;
        void getCurrentVersion().then(setVersion);
    }, []);

    const open = useCallback((path: string) => router.push(flowRouteFor(path)), [router]);

    const actions = [
        { key: "n", label: "New flow", run: () => setNewFlowOpen(true) },
        { key: "o", label: "Open", run: () => void openFlowFromPicker() },
        { key: "s", label: "Settings", run: () => setSettingsOpen(true) },
    ];

    // The cursor runs over the actions and then the recents as one column, so
    // j/k walks the whole screen the way it walks a buffer.
    const rows = entries ?? [];
    const total = actions.length + rows.length;

    useEffect(() => {
        function onKeyDown(e: KeyboardEvent) {
            if (e.altKey) return;
            // Meta/Ctrl chords are the OS's until proven otherwise; the two the
            // start screen claims are the ones the File menu also offers.
            if (e.metaKey || e.ctrlKey) {
                if (e.key === "n") {
                    e.preventDefault();
                    setNewFlowOpen(true);
                } else if (e.key === "o") {
                    e.preventDefault();
                    void openFlowFromPicker();
                }
                return;
            }
            const target = e.target as HTMLElement | null;
            if (target?.closest("input, textarea, [contenteditable='true'], [role='dialog']")) {
                return;
            }

            const action = actions.find((a) => a.key === e.key);
            if (action) {
                e.preventDefault();
                action.run();
                return;
            }
            if (/^[1-9]$/.test(e.key)) {
                const entry = rows[Number(e.key) - 1];
                if (entry) {
                    e.preventDefault();
                    open(entry.path);
                }
                return;
            }
            if (e.key === "?") {
                e.preventDefault();
                setCheatsheetOpen(true);
            } else if (e.key === "j" || e.key === "ArrowDown") {
                e.preventDefault();
                setCursor((c) => (total ? (c + 1) % total : 0));
            } else if (e.key === "k" || e.key === "ArrowUp") {
                e.preventDefault();
                setCursor((c) => (total ? (c - 1 + total) % total : 0));
            } else if (e.key === "Enter") {
                e.preventDefault();
                if (cursor < actions.length) actions[cursor].run();
                else open(rows[cursor - actions.length].path);
            }
        }

        window.addEventListener("keydown", onKeyDown);
        return () => window.removeEventListener("keydown", onKeyDown);
        // No dependency array on purpose: the handler closes over the cursor
        // and the recents, and re-subscribing on a screen this static is
        // cheaper than the refs it would take to avoid it.
    });

    return (
        <main
            className="flex min-h-screen items-center justify-center px-6 py-16"
            data-testid="start-screen"
        >
            <div className="w-full max-w-[34rem] font-mono text-sm">
                <div className="flex flex-col items-center gap-3">
                    <Wordmark animated className="h-10 w-auto" />
                    <div className="text-muted-foreground text-xs tracking-wide">
                        ebb{version && ` v${version}`}
                    </div>
                </div>

                <Rule />

                {actions.map((action, i) => (
                    <Row
                        key={action.key}
                        badge={action.key}
                        active={cursor === i}
                        onHover={() => setCursor(i)}
                        onSelect={action.run}
                        testid={`start-${action.label.split(" ")[0].toLowerCase()}`}
                    >
                        <span>{action.label}</span>
                    </Row>
                ))}

                <Rule />

                {entries === null ? (
                    <div className="space-y-2 px-2 py-1.5">
                        <Skeleton className="h-4 w-2/3" />
                        <Skeleton className="h-4 w-1/2" />
                    </div>
                ) : entries.length === 0 ? (
                    <p className="text-muted-foreground px-2 py-1.5 text-xs">
                        No flows yet. Press <Kbd>n</Kbd> to start one.
                    </p>
                ) : (
                    entries.map((entry, i) => (
                        <RecentRow
                            key={entry.path}
                            entry={entry}
                            index={i}
                            active={cursor === actions.length + i}
                            onHover={() => setCursor(actions.length + i)}
                            onSelect={() => open(entry.path)}
                        />
                    ))
                )}

                <Rule />

                <p className="text-muted-foreground flex flex-wrap justify-center gap-x-2 text-xs">
                    {LINKS.map((link, i) => (
                        <span key={link.href} className="whitespace-nowrap">
                            {i > 0 && <span aria-hidden="true">&middot; </span>}
                            {link.prefix}
                            <a
                                href={link.href}
                                target="_blank"
                                rel="noreferrer"
                                onClick={(e) => openExternal(e, link.href)}
                                className="hover:text-foreground underline underline-offset-2"
                            >
                                {link.label}
                            </a>
                        </span>
                    ))}
                </p>
            </div>
            <MigrationDialog onMigrated={refresh} />
        </main>
    );
}

interface RowProps {
    badge: string;
    active: boolean;
    onHover: () => void;
    onSelect: () => void;
    testid?: string;
    children: React.ReactNode;
}

function Row({ badge, active, onHover, onSelect, testid, children }: RowProps) {
    return (
        <button
            type="button"
            data-testid={testid}
            onMouseEnter={onHover}
            onClick={onSelect}
            className={cn(
                "flex w-full items-center gap-4 rounded px-2 py-1.5 text-left",
                active ? "bg-accent text-accent-foreground" : "",
            )}
        >
            <Kbd className="w-5 justify-center">{badge}</Kbd>
            <span className="min-w-0 flex-1">{children}</span>
        </button>
    );
}

interface RecentRowProps {
    entry: RecentEntry;
    index: number;
    active: boolean;
    onHover: () => void;
    onSelect: () => void;
}

function RecentRow({ entry, index, active, onHover, onSelect }: RecentRowProps) {
    return (
        <Row
            badge={String(index + 1)}
            active={active}
            onHover={onHover}
            onSelect={onSelect}
            testid={`start-recent-${index + 1}`}
        >
            <span className="flex items-baseline gap-2">
                <span className="truncate">{entry.label}</span>
                {entry.detail && (
                    <span className="text-muted-foreground truncate text-xs">{entry.detail}</span>
                )}
                <span className="text-muted-foreground ml-auto shrink-0 pl-2 text-xs">
                    {relativeTime(entry.updatedAt)}
                </span>
            </span>
            <span className="text-muted-foreground/70 block truncate text-xs">{entry.display}</span>
        </Row>
    );
}
