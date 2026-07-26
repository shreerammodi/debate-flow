import { render, screen, waitFor } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { beforeEach, describe, expect, it } from "vitest";

import FlowsFolderControl from "@/components/settings/FlowsFolderControl";
import { useFlowStore } from "@/lib/store/useFlowStore";

import { FLOWS_DIR, installFakeFlowFs, type FakeFlowFs } from "../../support/fakeFlowFs";

let fs: FakeFlowFs;

beforeEach(() => {
    fs = installFakeFlowFs();
    useFlowStore.setState({ flowsDir: null });
});

describe("FlowsFolderControl", () => {
    it("shows the resolved default when nothing is configured", async () => {
        render(<FlowsFolderControl />);
        // Shortened against home, since that is what the user recognizes.
        await waitFor(() =>
            expect(screen.getByTestId("flows-folder-path")).toHaveTextContent("~/Documents/ebb"),
        );
    });

    it("shows the configured folder instead once one is set", async () => {
        useFlowStore.setState({ flowsDir: "/Volumes/usb/rounds" });
        render(<FlowsFolderControl />);
        await waitFor(() =>
            expect(screen.getByTestId("flows-folder-path")).toHaveTextContent(
                "/Volumes/usb/rounds",
            ),
        );
    });

    it("stores the folder the picker returns", async () => {
        fs.nextDirectory = "/Volumes/usb/rounds";
        render(<FlowsFolderControl />);

        await userEvent.click(screen.getByTestId("flows-folder-choose"));

        await waitFor(() => expect(useFlowStore.getState().flowsDir).toBe("/Volumes/usb/rounds"));
    });

    it("keeps the current folder when the picker is cancelled", async () => {
        useFlowStore.setState({ flowsDir: "/Volumes/usb/rounds" });
        fs.nextDirectory = null;
        render(<FlowsFolderControl />);

        await userEvent.click(screen.getByTestId("flows-folder-choose"));

        expect(useFlowStore.getState().flowsDir).toBe("/Volumes/usb/rounds");
    });

    it("clears the override rather than pinning today's default", async () => {
        useFlowStore.setState({ flowsDir: "/Volumes/usb/rounds" });
        render(<FlowsFolderControl />);

        await userEvent.click(screen.getByTestId("flows-folder-reset"));

        expect(useFlowStore.getState().flowsDir).toBeNull();
        await waitFor(() =>
            expect(screen.getByTestId("flows-folder-path")).toHaveTextContent("~/Documents/ebb"),
        );
    });

    it("offers no reset while following the default", async () => {
        render(<FlowsFolderControl />);
        await waitFor(() =>
            expect(screen.getByTestId("flows-folder-path")).not.toBeEmptyDOMElement(),
        );
        expect(screen.queryByTestId("flows-folder-reset")).not.toBeInTheDocument();
        expect(FLOWS_DIR).toContain("Documents/ebb");
    });
});
