import { render, screen } from "@testing-library/react";
import { act } from "react";
import { beforeEach, describe, expect, it } from "vitest";

import ShareSheet from "@/components/collab/ShareSheet";
import {
    closeShareSheet,
    openShareSheet,
    showShareCode,
    showShareFailure,
    showShareGuest,
} from "@/lib/store/useShareSheet";

beforeEach(() => {
    closeShareSheet();
});

describe("ShareSheet", () => {
    it("renders nothing until a share opens it", () => {
        render(<ShareSheet />);
        expect(screen.queryByTestId("share-sheet")).toBeNull();
    });

    it("shows Getting ready, and no code", () => {
        render(<ShareSheet />);
        act(() => openShareSheet("editor", ""));
        expect(screen.getByTestId("share-status").textContent).toBe("Getting ready...");
        expect(screen.queryByTestId("share-code")).toBeNull();
    });

    it("shows the code in two groups, and who it is waiting for", () => {
        render(<ShareSheet />);
        act(() => {
            openShareSheet("editor", "");
            showShareCode("K7QM3XPV", async () => {});
        });
        expect(screen.getByTestId("share-code").textContent).toBe("K7QM-3XPV");
        expect(screen.getByTestId("share-status").textContent).toContain(
            "Waiting for your partner",
        );
    });

    it("says a guest arrived", () => {
        render(<ShareSheet />);
        act(() => {
            openShareSheet("editor", "");
            showShareCode("K7QM3XPV", async () => {});
            showShareGuest("Sam");
        });
        expect(screen.getByTestId("share-status").textContent).toBe("Sam joined");
    });

    it("names a view-only code for what it grants", () => {
        render(<ShareSheet />);
        act(() => openShareSheet("viewer", ""));
        expect(screen.getByTestId("share-sheet").textContent).toContain("Share view only");
    });

    it("shows the warning before there is a code to warn about", () => {
        render(<ShareSheet />);
        act(() =>
            openShareSheet("editor", "Relaying is off, so this code only works on the same wifi."),
        );
        expect(screen.getByTestId("share-warning").textContent).toBe(
            "Relaying is off, so this code only works on the same wifi.",
        );
    });

    it("shows the reason, and no code, when there is none", () => {
        render(<ShareSheet />);
        act(() => {
            openShareSheet("editor", "");
            showShareFailure("Could not reach the relay for that code");
        });
        expect(screen.getByTestId("share-status").textContent).toBe(
            "Could not reach the relay for that code",
        );
        expect(screen.queryByTestId("share-code")).toBeNull();
    });
});
