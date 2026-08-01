"use client";

import { Eye, ShareNetwork, SignIn, SignOut, UserPlus } from "@phosphor-icons/react";

import SettingRow from "@/components/settings/SettingRow";
import { Button } from "@/components/ui/button";
import { executeCommand } from "@/lib/commands/commands";
import { useCollabStore } from "@/lib/store/useCollabStore";
import { useFlowStore } from "@/lib/store/useFlowStore";

/**
 * The visible way into a session, for a debater who has not learned the
 * palette. Every button runs the same command the palette does, so there is
 * still one route to a session and the master switch still guards it.
 *
 * A ticket is minted, not exchanged with a service: sharing puts one on screen
 * to send however the two of you already talk, and joining takes one back,
 * which is why the two buttons are asymmetric.
 */
export default function SessionControls() {
    const status = useCollabStore((s) => s.status);
    const round = useFlowStore((s) => s.round);
    const contacts = useFlowStore((s) => s.contacts);

    const live = status !== "off";
    const hasContacts = Object.keys(contacts).length > 0;

    return (
        <SettingRow
            title="Session"
            description={
                live
                    ? "A session is running. Showing an invite again creates a new one."
                    : "Sharing starts a session and creates an invite a partner can edit or watch your flow with. Each invite works once, and only allows access to this round."
            }
        >
            <div className="flex flex-wrap gap-2" data-testid="session-controls">
                <Button
                    type="button"
                    size="sm"
                    disabled={!round}
                    onClick={() => executeCommand("collab.share")}
                    data-testid="session-share"
                >
                    <ShareNetwork />
                    {live ? "Show invite" : "Share this round"}
                </Button>
                <Button
                    type="button"
                    size="sm"
                    variant="outline"
                    disabled={!round}
                    onClick={() => executeCommand("collab.shareView")}
                    data-testid="session-share-view"
                >
                    <Eye />
                    Share view only
                </Button>
                <Button
                    type="button"
                    size="sm"
                    variant="outline"
                    disabled={!round || !hasContacts}
                    onClick={() => executeCommand("collab.invite")}
                    data-testid="session-invite"
                >
                    <UserPlus />
                    Invite a partner
                </Button>
                <Button
                    type="button"
                    size="sm"
                    variant="outline"
                    onClick={() => executeCommand("collab.join")}
                    data-testid="session-join"
                >
                    <SignIn />
                    Join with an invite
                </Button>
                {live && (
                    <Button
                        type="button"
                        size="sm"
                        variant="outline"
                        className="text-warn"
                        onClick={() => executeCommand("collab.end")}
                        data-testid="session-end"
                    >
                        <SignOut />
                        End session
                    </Button>
                )}
            </div>
        </SettingRow>
    );
}
