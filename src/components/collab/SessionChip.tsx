"use client";

import { useId, useState } from "react";

import { Button } from "@/components/ui/button";
import { disconnectPeer, endSession } from "@/lib/collab/runtime";
import { type CollabPeerView, type CollabStatus, useCollabStore } from "@/lib/store/useCollabStore";
import { cn } from "@/lib/utils";

type LiveStatus = Exclude<CollabStatus, "off">;

const STATUS_LABEL: Record<LiveStatus, string> = {
    connecting: "Connecting",
    connected: "Connected",
    reconnecting: "Reconnecting",
};

const STATUS_DOT: Record<LiveStatus, string> = {
    connecting: "bg-warn",
    connected: "bg-good",
    reconnecting: "bg-warn",
};

/** What each role may do, in the words the chip shows a debater. */
const ROLE_LABEL: Record<CollabPeerView["role"], string> = {
    partner: "edit",
    coach: "view",
};

function peerCountLabel(count: number): string {
    if (count === 0) return "no peers";
    return count === 1 ? "1 peer" : `${count} peers`;
}

export interface SessionChipProps {
    className?: string;
}

/**
 * The bottom-left session chip: connection state and peer count, expanding on
 * click into one row per peer with role and connection type. A side the host
 * admitted as a coach is told so, because a grid that refuses every keystroke
 * needs to say why.
 *
 * A session is the only thing it reports, so `status: "off"` renders nothing at
 * all - the master switch leaves no trace in the DOM. It is a plain button and
 * a plain panel, never a dialog: nothing here blocks the grid, autofocuses, or
 * traps focus, because a debater mid-speech cannot afford either.
 */
export default function SessionChip({ className }: SessionChipProps) {
    const status = useCollabStore((s) => s.status);
    const peers = useCollabStore((s) => s.peers);
    const selfRole = useCollabStore((s) => s.selfRole);
    const [expanded, setExpanded] = useState(false);
    const panelId = useId();

    // A session that ends closes the panel, so the next one opens collapsed.
    // Adjusted during render rather than in an effect: an effect would take a
    // second pass to settle, and React documents this shape for exactly the
    // case of resetting state when something outside the component changes.
    if (status === "off" && expanded) setExpanded(false);
    if (status === "off") return null;

    return (
        // The caller owns the outer slot's position; the anchor stays `relative`
        // underneath it, so a `fixed` slot class has nothing to fight with.
        <div className={className}>
            <div className="relative">
                {/* The trigger leads in DOM order so a Tab from the chip walks
                    into the panel, while `bottom-full` draws the panel above. */}
                <button
                    type="button"
                    data-testid="collab-chip"
                    data-state={status}
                    aria-expanded={expanded}
                    aria-controls={expanded ? panelId : undefined}
                    onClick={() => setExpanded((open) => !open)}
                    className={cn(
                        "border-border bg-card text-foreground flex w-full items-center gap-1.5 rounded-full border",
                        "hover:bg-accent px-2.5 py-1 text-[12px] transition-colors focus-visible:outline-2",
                    )}
                >
                    <span
                        aria-hidden="true"
                        className={cn("size-1.5 shrink-0 rounded-full", STATUS_DOT[status])}
                    />
                    <span className="font-medium">{STATUS_LABEL[status]}</span>
                    <span className="text-muted-foreground truncate">
                        {peerCountLabel(peers.length)}
                    </span>
                </button>
                {expanded && (
                    <div
                        id={panelId}
                        data-testid="collab-chip-peers"
                        className="border-border bg-card absolute bottom-full left-0 z-30 mb-1 flex w-48 flex-col gap-1 rounded-md border p-1.5 shadow-md"
                    >
                        {selfRole === "coach" && (
                            <p
                                data-testid="collab-self-role"
                                className="text-muted-foreground border-border border-b px-1 py-0.5 text-[12px]"
                            >
                                You are viewing this round, not editing it.
                            </p>
                        )}
                        {peers.length === 0 ? (
                            <p className="text-muted-foreground px-1 py-0.5 text-[12px]">
                                No peers connected.
                            </p>
                        ) : (
                            peers.map((peer) => (
                                <div
                                    key={peer.endpointId}
                                    data-testid="collab-peer-row"
                                    className="flex flex-col gap-0.5 rounded px-1 py-0.5"
                                >
                                    <div className="flex items-center gap-1.5">
                                        <span className="min-w-0 flex-1 truncate text-[12px] font-medium">
                                            {peer.name}
                                        </span>
                                        <Button
                                            type="button"
                                            variant="ghost"
                                            size="xs"
                                            data-testid="collab-peer-disconnect"
                                            aria-label={`Disconnect ${peer.name}`}
                                            onClick={() => void disconnectPeer(peer.endpointId)}
                                        >
                                            Disconnect
                                        </Button>
                                    </div>
                                    <div className="text-muted-foreground flex items-center gap-1.5 text-[10px]">
                                        <span
                                            data-testid="collab-peer-role"
                                            className="border-border shrink-0 rounded-full border px-1.5 py-px"
                                        >
                                            {ROLE_LABEL[peer.role]}
                                        </span>
                                        <span
                                            data-testid="collab-peer-connection"
                                            className="shrink-0"
                                        >
                                            {peer.connectionType}
                                        </span>
                                    </div>
                                </div>
                            ))
                        )}
                        <Button
                            type="button"
                            variant="ghost"
                            size="xs"
                            className="text-warn justify-start"
                            data-testid="collab-end-session"
                            onClick={() => void endSession()}
                        >
                            End session
                        </Button>
                    </div>
                )}
            </div>
        </div>
    );
}
