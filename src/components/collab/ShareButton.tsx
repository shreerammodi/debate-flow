"use client";

import { Eye, ShareNetwork, SignIn, UserPlus } from "@phosphor-icons/react";

import { Button } from "@/components/ui/button";
import { Tip } from "@/components/ui/tooltip";
import { executeCommand } from "@/lib/commands/commands";
import { useFlowStore } from "@/lib/store/useFlowStore";
import { isDesktop } from "@/lib/update/adapter";

/**
 * Sharing, beside the round it shares.
 *
 * Not in Settings: Settings is where a debater changes how the application
 * behaves, and this does a thing to the flow that is open. Every button runs
 * the same command the palette does, so there is still one route to a session
 * and the consent question still guards it.
 *
 * Absent off the desktop, the way the Collaboration settings pane is. A
 * session is an iroh endpoint, which a browser cannot bind, so a button here
 * would offer a debater something that cannot exist and answer their click
 * with nothing.
 */
export default function ShareButton() {
    const round = useFlowStore((s) => s.round);
    const contacts = useFlowStore((s) => s.contacts);
    const hasContacts = Object.keys(contacts).length > 0;

    if (!isDesktop()) return null;

    return (
        <div className="flex flex-wrap items-center gap-1" data-testid="share-controls">
            {round && (
                <>
                    <Button
                        type="button"
                        size="xs"
                        variant="outline"
                        data-testid="sidebar-share"
                        onClick={() => executeCommand("collab.share")}
                    >
                        <ShareNetwork />
                        Invite partner
                    </Button>
                    <Tip label="Share view only" command="collab.shareView">
                        <Button
                            type="button"
                            size="xs"
                            variant="ghost"
                            aria-label="Share view only"
                            data-testid="sidebar-share-view"
                            onClick={() => executeCommand("collab.shareView")}
                        >
                            <Eye />
                        </Button>
                    </Tip>
                    {hasContacts && (
                        <Tip label="Invite a saved partner" command="collab.invite">
                            <Button
                                type="button"
                                size="xs"
                                variant="ghost"
                                aria-label="Invite a saved partner"
                                data-testid="sidebar-invite"
                                onClick={() => executeCommand("collab.invite")}
                            >
                                <UserPlus />
                            </Button>
                        </Tip>
                    )}
                </>
            )}
            <Tip label="Join with a code" command="collab.join">
                <Button
                    type="button"
                    size="xs"
                    variant="ghost"
                    aria-label="Join with a code"
                    data-testid="sidebar-join"
                    onClick={() => executeCommand("collab.join")}
                >
                    <SignIn />
                </Button>
            </Tip>
        </div>
    );
}
