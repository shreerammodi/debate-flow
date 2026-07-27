"use client";

import SettingRow from "@/components/settings/SettingRow";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import {
    Select,
    SelectContent,
    SelectItem,
    SelectTrigger,
    SelectValue,
} from "@/components/ui/select";
import { addContact, type Contacts, removeContact } from "@/lib/collab/contacts";
import type { Role } from "@/lib/collab/types";
import { useFlowStore } from "@/lib/store/useFlowStore";

/** What each role may do, in the words the chip shows a debater. */
const ROLE_OPTIONS: { value: Role; label: string }[] = [
    { value: "partner", label: "can edit" },
    { value: "coach", label: "view only" },
];

/** An EndpointId is 52 characters of base32; a row shows the first eight. */
const SHORT_ID = 8;

/**
 * The saved peers, edited in place. Entries arrive from one click on a toast
 * after a session, so the list renames, re-roles, and drops them; there is no
 * way to type an EndpointId in by hand, which is the point of a contact.
 */
export default function ContactList() {
    const contacts = useFlowStore((s) => s.contacts);
    const setContacts = useFlowStore((s) => s.setContacts);

    // Insertion order, so a rename never moves the row out from under the
    // cursor typing it.
    const entries = Object.entries(contacts);

    return (
        <SettingRow
            title="Contacts"
            description="Peers you have shared with. Inviting one dials them by name, with no ticket."
        >
            {entries.length === 0 ? (
                <p className="text-muted-foreground text-[12px]" data-testid="contact-list-empty">
                    No partners saved yet. After a session, saving one is a click on the toast.
                </p>
            ) : (
                <ul className="m-0 flex list-none flex-col gap-1.5 p-0">
                    {entries.map(([endpointId, contact]) => {
                        const short = endpointId.slice(0, SHORT_ID);
                        const shown = contact.name.trim() === "" ? short : contact.name;
                        return (
                            <li
                                key={endpointId}
                                data-testid={`contact-row-${endpointId}`}
                                className="flex items-center gap-2"
                            >
                                <Input
                                    value={contact.name}
                                    onChange={(e) =>
                                        setContacts(
                                            addContact(contacts, endpointId, {
                                                ...contact,
                                                name: e.target.value,
                                            }),
                                        )
                                    }
                                    // A nameless entry is unreadable to the config
                                    // parser, so an emptied field falls back to the
                                    // short id rather than losing the peer.
                                    onBlur={() => {
                                        if (contact.name.trim() !== "") return;
                                        setContacts(
                                            addContact(contacts, endpointId, {
                                                ...contact,
                                                name: short,
                                            }),
                                        );
                                    }}
                                    aria-label={`Name for ${shown}`}
                                    data-testid={`contact-name-${endpointId}`}
                                    className="h-8 min-w-0 flex-1"
                                />
                                <Select
                                    value={contact.role}
                                    // Base UI Select renders the raw value unless given a
                                    // value->label map to resolve the trigger display.
                                    items={ROLE_OPTIONS}
                                    onValueChange={(value) =>
                                        setContacts(
                                            addContact(contacts, endpointId, {
                                                ...contact,
                                                role: value as Role,
                                            }),
                                        )
                                    }
                                >
                                    <SelectTrigger
                                        size="sm"
                                        aria-label={`Role for ${shown}`}
                                        data-testid={`contact-role-${endpointId}`}
                                        className="w-28 shrink-0"
                                    >
                                        <SelectValue />
                                    </SelectTrigger>
                                    <SelectContent>
                                        {ROLE_OPTIONS.map((r) => (
                                            <SelectItem
                                                key={r.value}
                                                value={r.value}
                                                data-testid={`contact-role-${endpointId}-${r.value}`}
                                            >
                                                {r.label}
                                            </SelectItem>
                                        ))}
                                    </SelectContent>
                                </Select>
                                <span
                                    title={endpointId}
                                    className="text-muted-foreground shrink-0 font-mono text-[11px]"
                                >
                                    {short}
                                </span>
                                <Button
                                    type="button"
                                    variant="ghost"
                                    size="xs"
                                    aria-label={`Remove ${shown}`}
                                    data-testid={`contact-remove-${endpointId}`}
                                    onClick={() => setContacts(removeContact(contacts, endpointId))}
                                >
                                    Remove
                                </Button>
                            </li>
                        );
                    })}
                </ul>
            )}
        </SettingRow>
    );
}
