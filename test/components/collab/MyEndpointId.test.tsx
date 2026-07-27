/**
 * MyEndpointId component tests.
 *
 * The row exists so a partner can be handed an identity before either side has
 * a round, so what it shows before an endpoint is bound matters as much as
 * what it shows after.
 */

import { render, screen } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { act } from "react";
import { beforeEach, describe, expect, it, vi } from "vitest";

import MyEndpointId from "@/components/collab/MyEndpointId";
import { useCollabStore } from "@/lib/store/useCollabStore";

const { toastSuccess, toastError } = vi.hoisted(() => ({
    toastSuccess: vi.fn(),
    toastError: vi.fn(),
}));

vi.mock("sonner", () => ({ toast: { success: toastSuccess, error: toastError } }));

const ID = "c".repeat(64);

beforeEach(() => {
    toastSuccess.mockClear();
    toastError.mockClear();
    useCollabStore.setState({ endpointId: null });
});

describe("MyEndpointId", () => {
    it("has nothing to copy until an endpoint has bound one", () => {
        render(<MyEndpointId />);
        expect(screen.getByTestId("my-id-copy")).toBeDisabled();
        expect(screen.getByTestId("my-id").textContent).not.toContain("c");
    });

    it("shows the id once the listener reports it", () => {
        render(<MyEndpointId />);
        act(() => useCollabStore.getState().setEndpointId(ID));

        expect(screen.getByTestId("my-id").textContent).toBe(ID);
        expect(screen.getByTestId("my-id-copy")).not.toBeDisabled();
    });

    it("copies the whole id, not the short form a row shows", async () => {
        const writeText = vi.fn(async () => {});
        vi.stubGlobal("navigator", { ...navigator, clipboard: { writeText } });
        useCollabStore.setState({ endpointId: ID });
        render(<MyEndpointId />);

        await userEvent.click(screen.getByTestId("my-id-copy"));
        expect(writeText).toHaveBeenCalledWith(ID);
        expect(toastSuccess).toHaveBeenCalled();
        vi.unstubAllGlobals();
    });

    it("selects the id when the webview refuses the write", async () => {
        const writeText = vi.fn(async () => {
            throw new Error("The request is not allowed by the user agent");
        });
        vi.stubGlobal("navigator", { ...navigator, clipboard: { writeText } });
        useCollabStore.setState({ endpointId: ID });
        render(<MyEndpointId />);

        await userEvent.click(screen.getByTestId("my-id-copy"));
        expect(toastError.mock.calls[0]?.[0]).toMatch(/Cmd\+C/);
        expect(window.getSelection()?.toString()).toBe(ID);
        vi.unstubAllGlobals();
    });
});
