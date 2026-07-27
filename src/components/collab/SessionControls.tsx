"use client";

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
 * A ticket is minted, not exchanged with a service: sharing copies one to the
 * clipboard and joining reads one back, which is why the two buttons are
 * asymmetric.
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
                    ? "A session is running. Copying the invite again mints a fresh one."
                    : "Sharing starts a session and copies an invite. It works once, and it carries this round only."
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
                    {live ? "Copy invite" : "Share this round"}
                </Button>
                <Button
                    type="button"
                    size="sm"
                    variant="outline"
                    disabled={!round || !hasContacts}
                    onClick={() => executeCommand("collab.invite")}
                    data-testid="session-invite"
                >
                    Invite a partner
                </Button>
                <Button
                    type="button"
                    size="sm"
                    variant="outline"
                    onClick={() => executeCommand("collab.join")}
                    data-testid="session-join"
                >
                    Join from clipboard
                </Button>
                {live && (
                    <Button
                        type="button"
                        size="sm"
                        variant="ghost"
                        onClick={() => executeCommand("collab.end")}
                        data-testid="session-end"
                    >
                        End session
                    </Button>
                )}
            </div>
        </SettingRow>
    );
}
