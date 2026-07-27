/**
 * DisplayNameRow component tests.
 *
 * The hostname is shown as a placeholder and never as a value: a value would
 * be saved, and the config file it is saved to syncs between machines.
 */

import { render, screen, waitFor } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { beforeEach, describe, expect, it, vi } from "vitest";

import DisplayNameRow from "@/components/collab/DisplayNameRow";
import { useFlowStore } from "@/lib/store/useFlowStore";

const machineName = vi.hoisted(() => vi.fn(async () => "smodi-mbp"));

vi.mock("@/lib/collab/machineName", () => ({ machineName }));

beforeEach(() => {
    window.localStorage.clear();
    machineName.mockClear();
    useFlowStore.setState({ collabName: "" });
});

describe("DisplayNameRow", () => {
    it("offers the machine's name as a placeholder, leaving the setting empty", async () => {
        render(<DisplayNameRow />);

        await waitFor(() =>
            expect(screen.getByTestId("collab-name")).toHaveAttribute("placeholder", "smodi-mbp"),
        );
        expect(screen.getByTestId("collab-name")).toHaveValue("");
        expect(useFlowStore.getState().collabName).toBe("");
    });

    it("keeps the field usable when the shell has no name to give", async () => {
        machineName.mockResolvedValueOnce("");
        render(<DisplayNameRow />);

        await waitFor(() =>
            expect(screen.getByTestId("collab-name")).toHaveAttribute(
                "placeholder",
                "This machine",
            ),
        );
    });

    it("saves what the debater types", async () => {
        render(<DisplayNameRow />);

        await userEvent.type(screen.getByTestId("collab-name"), "Rin");
        expect(useFlowStore.getState().collabName).toBe("Rin");
    });

    it("shows a name already set instead of the placeholder", () => {
        useFlowStore.setState({ collabName: "Rin" });
        render(<DisplayNameRow />);

        expect(screen.getByTestId("collab-name")).toHaveValue("Rin");
    });
});
