import { describe, expect, it } from "vitest";

import { chipSummary, pendingLine } from "@/lib/collab/statusLine";

describe("chipSummary", () => {
    it("names the one partner who is there", () => {
        expect(chipSummary([{ name: "Sam", relayed: false }], [])).toBe("Connected to Sam");
    });

    it("says a link is relayed without making it an alarm", () => {
        expect(chipSummary([{ name: "Sam", relayed: true }], [])).toBe("Connected to Sam, relayed");
    });

    it("counts rather than lists once there is more than one", () => {
        expect(
            chipSummary(
                [
                    { name: "Sam", relayed: false },
                    { name: "Kim", relayed: true },
                ],
                [],
            ),
        ).toBe("Connected to 2 partners");
    });

    it("names the partner it is waiting for", () => {
        expect(chipSummary([], [{ name: "Sam", unreachable: false }])).toBe("Waiting for Sam");
    });

    it("says so when it cannot reach the partner it is waiting for", () => {
        expect(chipSummary([], [{ name: "Sam", unreachable: true }])).toBe("Can't reach Sam");
    });

    it("counts partners it is waiting for once there is more than one", () => {
        expect(
            chipSummary(
                [],
                [
                    { name: "Sam", unreachable: false },
                    { name: "Kim", unreachable: true },
                ],
            ),
        ).toBe("Waiting for 2 partners");
    });

    it("leads with who is here when some are and some are not", () => {
        expect(
            chipSummary([{ name: "Sam", relayed: false }], [{ name: "Kim", unreachable: true }]),
        ).toBe("Connected to Sam, waiting for 1 more");
    });

    it("says a session with nobody in it is open, not broken", () => {
        expect(chipSummary([], [])).toBe("Waiting to be joined");
    });
});

describe("pendingLine", () => {
    it("says a partner has not opened the round yet", () => {
        expect(pendingLine({ name: "Sam", unreachable: false })).toBe(
            "Waiting for Sam to open this round",
        );
    });

    it("says what to do about a partner it cannot reach", () => {
        expect(pendingLine({ name: "Sam", unreachable: true })).toBe(
            "Can't reach Sam. You both need internet, or the same wifi.",
        );
    });
});
