/**
 * Debate event definitions. An event lists each side's speeches in speaking
 * order; the full column order for a round is derived by strictly
 * alternating the two lists starting with the first-speaking side
 * (speechOrder). Policy fixes the aff as first speaker; PF's first speaker
 * comes from the flip (FlowRound.firstSide).
 */

import type { Side } from "@/lib/model/types";

export type EventId = "policy" | "pf" | "ld" | "parli";

export interface SpeechDef {
    id: string;
    name: string;
    /** Column-header label; equals name for Policy. */
    short: string;
    side: Side;
}

export interface CrossExPeriod {
    label: string;
    /** Which team holds the question column: the first- or second-speaking side. */
    q: "first" | "second";
}

/** What an event calls one of its sides, and who debates on it. */
export interface SideLabel {
    /** Short side name, as headers, buttons, and exports print it. */
    label: string;
    /** The side's debater slots, in the order they speak. */
    speakers: [string, string];
}

export interface EventDef {
    id: EventId;
    name: string;
    /** Each side's speeches in that side's own speaking order. */
    aff: SpeechDef[];
    neg: SpeechDef[];
    /** The flip decides who speaks first (PF); false = always aff-first. */
    variableOrder: boolean;
    /** `shared`: both sides question each other (PF crossfire), so columns are
     *  labelled by side rather than as questioner/responder. Absent for events
     *  with no cross-examination at all, which get no cross-ex sheet. */
    crossEx?: { title: string; periods: CrossExPeriod[]; shared?: boolean };
    /** Side naming, when the event does not call its sides aff and neg. */
    sides?: Record<Side, SideLabel>;
}

const AFF_NEG_SIDES: Record<Side, SideLabel> = {
    aff: { label: "Aff", speakers: ["1A", "2A"] },
    neg: { label: "Neg", speakers: ["1N", "2N"] },
};

const speech = (id: string, name: string, short: string, side: Side): SpeechDef => ({
    id,
    name,
    short,
    side,
});

export const EVENTS: Record<EventId, EventDef> = {
    policy: {
        id: "policy",
        name: "Policy",
        aff: [
            speech("1ac", "1AC", "1AC", "aff"),
            speech("2ac", "2AC", "2AC", "aff"),
            speech("1ar", "1AR", "1AR", "aff"),
            speech("2ar", "2AR", "2AR", "aff"),
        ],
        neg: [
            speech("1nc", "1NC", "1NC", "neg"),
            speech("block", "Block", "Block", "neg"),
            speech("2nr", "2NR", "2NR", "neg"),
        ],
        variableOrder: false,
        crossEx: {
            title: "CX",
            periods: [
                { label: "1AC CX", q: "second" },
                { label: "1NC CX", q: "first" },
                { label: "2AC CX", q: "second" },
                { label: "2NC CX", q: "first" },
            ],
        },
    },
    pf: {
        id: "pf",
        name: "Public Forum",
        aff: [
            speech("ac", "Aff Constructive", "AC", "aff"),
            speech("ar", "Aff Rebuttal", "AR", "aff"),
            speech("as", "Aff Summary", "AS", "aff"),
            speech("af", "Aff Final Focus", "AF", "aff"),
        ],
        neg: [
            speech("nc", "Neg Constructive", "NC", "neg"),
            speech("nr", "Neg Rebuttal", "NR", "neg"),
            speech("ns", "Neg Summary", "NS", "neg"),
            speech("nf", "Neg Final Focus", "NF", "neg"),
        ],
        variableOrder: true,
        crossEx: {
            title: "Cross-Examination",
            shared: true,
            periods: [
                { label: "First Cross", q: "first" },
                { label: "Second Cross", q: "first" },
                { label: "Grand Cross", q: "first" },
            ],
        },
    },
    ld: {
        id: "ld",
        name: "Lincoln-Douglas",
        aff: [
            speech("1ac", "1AC", "1AC", "aff"),
            speech("1ar", "1AR", "1AR", "aff"),
            speech("2ar", "2AR", "2AR", "aff"),
        ],
        neg: [speech("1nc", "1NC", "1NC", "neg"), speech("2nr", "2NR", "2NR", "neg")],
        variableOrder: false,
        crossEx: {
            title: "CX",
            periods: [
                { label: "1AC CX", q: "second" },
                { label: "1NC CX", q: "first" },
            ],
        },
    },
    parli: {
        id: "parli",
        name: "Parliamentary",
        // The opening speech is the Prime Minister, not a "PMC"; only the
        // speeches that have a rebuttal counterpart are named Constructive.
        aff: [
            speech("pm", "Prime Minister", "PM", "aff"),
            speech("mgc", "Member of the Government Constructive", "MGC", "aff"),
            speech("pmr", "Prime Minister Rebuttal", "PMR", "aff"),
        ],
        // The MOC and LOR run back to back, so they share one column the way
        // Policy's 2NC and 1NR share the Block: strict alternation cannot
        // express two consecutive speeches on the same side.
        neg: [
            speech("loc", "Leader of the Opposition Constructive", "LOC", "neg"),
            speech("block", "Opposition Block", "Block", "neg"),
        ],
        variableOrder: false,
        // Parliamentary has no cross-examination; a point of information
        // interrupts a speech rather than occupying a period of its own.
        sides: {
            aff: { label: "Gov", speakers: ["PM", "MG"] },
            neg: { label: "Opp", speakers: ["LO", "MO"] },
        },
    },
};

export function getEvent(id?: EventId): EventDef {
    return EVENTS[id ?? "policy"];
}

/** The event's side naming, falling back to the aff/neg the model stores. */
export function sideLabels(id?: EventId): Record<Side, SideLabel> {
    return getEvent(id).sides ?? AFF_NEG_SIDES;
}

/**
 * The round's full column order: the two side lists strictly alternated,
 * starting with firstSide. Uneven lists (Policy: 4 aff / 3 neg) interleave
 * until the shorter runs out, then the longer's tail follows.
 */
export function speechOrder(event: EventDef, firstSide: Side): SpeechDef[] {
    const first = firstSide === "aff" ? event.aff : event.neg;
    const second = firstSide === "aff" ? event.neg : event.aff;
    const order: SpeechDef[] = [];
    for (let i = 0; i < Math.max(first.length, second.length); i++) {
        if (first[i]) order.push(first[i]);
        if (second[i]) order.push(second[i]);
    }
    return order;
}
