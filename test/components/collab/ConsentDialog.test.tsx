import { render, screen } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { act } from "react";
import { beforeEach, describe, expect, it } from "vitest";

import ConsentDialog from "@/components/collab/ConsentDialog";
import { useCollabConsent } from "@/lib/store/useCollabConsent";

beforeEach(() => {
    useCollabConsent.getState().close();
});

describe("ConsentDialog", () => {
    it("renders nothing until something asks", () => {
        render(<ConsentDialog />);
        expect(screen.queryByTestId("collab-consent")).toBeNull();
    });

    it("says what sharing does and what it does not do yet", () => {
        render(<ConsentDialog />);
        act(() => useCollabConsent.getState().ask());
        expect(screen.getByTestId("collab-consent").textContent).toContain(
            "Sharing lets ebb connect to your partner over the network. Nothing is sent anywhere until you share a round.",
        );
    });

    it("answers yes on Turn on sharing", async () => {
        render(<ConsentDialog />);
        act(() => useCollabConsent.getState().ask());
        await userEvent.click(screen.getByTestId("collab-consent-yes"));
        expect(useCollabConsent.getState().open).toBe(false);
    });

    it("answers no on Not now", async () => {
        render(<ConsentDialog />);
        act(() => useCollabConsent.getState().ask());
        await userEvent.click(screen.getByTestId("collab-consent-no"));
        expect(useCollabConsent.getState().open).toBe(false);
    });

    // The question exists for a debater who did not mean to click Share, so
    // the answer their reflexes give has to be the one that reaches nobody.
    it("holds the focus on Not now, and takes Escape as one", async () => {
        render(<ConsentDialog />);
        act(() => useCollabConsent.getState().ask());
        expect(screen.getByTestId("collab-consent-no")).toHaveFocus();
        await userEvent.keyboard("{Escape}");
        expect(useCollabConsent.getState().open).toBe(false);
    });
});
