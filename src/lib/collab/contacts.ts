/**
 * The peers a debater has shared with before.
 *
 * A contact exists so nobody copies a 52-character key by hand: entries are
 * created by one click on a toast after a session that worked, and an
 * EndpointId is stable per install, so the same partner is reachable the next
 * time with no ticket at all.
 *
 * The table is keyed by EndpointId and lives in the config file, which is
 * hand-editable and synced between machines. So parsing is total: anything
 * unrecognizable degrades to absent rather than to a half-valid contact.
 */

import type { Role } from "./types";

export interface Contact {
    name: string;
    role: Role;
}

/** EndpointId to contact. */
export type Contacts = Record<string, Contact>;

/** How much of an EndpointId is worth showing when there is no name. */
const SHORT_ID = 8;

function isRole(value: unknown): value is Role {
    return value === "partner" || value === "coach";
}

/**
 * A contact table from whatever the config file held.
 *
 * An unknown role is dropped rather than defaulted: defaulting would decide,
 * from a typo, whether that peer may write into the round.
 */
export function resolveContacts(raw: unknown): Contacts {
    if (raw === null || typeof raw !== "object" || Array.isArray(raw)) return {};
    const out: Contacts = {};
    for (const [endpointId, value] of Object.entries(raw as Record<string, unknown>)) {
        if (!endpointId) continue;
        if (value === null || typeof value !== "object") continue;
        const entry = value as Partial<Contact>;
        if (typeof entry.name !== "string" || entry.name.trim() === "") continue;
        if (!isRole(entry.role)) continue;
        out[endpointId] = { name: entry.name, role: entry.role };
    }
    return out;
}

export function addContact(contacts: Contacts, endpointId: string, contact: Contact): Contacts {
    return { ...contacts, [endpointId]: contact };
}

export function removeContact(contacts: Contacts, endpointId: string): Contacts {
    if (!(endpointId in contacts)) return contacts;
    const { [endpointId]: _gone, ...rest } = contacts;
    return rest;
}

/** What to call this peer on screen. */
export function contactName(contacts: Contacts, endpointId: string): string {
    return contacts[endpointId]?.name ?? endpointId.slice(0, SHORT_ID);
}

export function isKnown(contacts: Contacts, endpointId: string): boolean {
    return endpointId in contacts;
}
