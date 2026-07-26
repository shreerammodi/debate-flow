"use client";

import { useRouter } from "next/navigation";
import { useEffect } from "react";

import { flowRouteFor, setFlowNavigator } from "@/lib/commands/flowNav";

/**
 * Registers the live router for commands that run outside React - the keyboard
 * layer and the native menu both dispatch from module scope. Renders nothing;
 * mounted once in the root layout beside the other app-wide hosts.
 */
export default function NavigatorHost() {
    const router = useRouter();

    useEffect(() => {
        setFlowNavigator({
            openPath: (path) => router.push(flowRouteFor(path)),
            toStart: () => router.push("/"),
        });
        return () => setFlowNavigator(null);
    }, [router]);

    return null;
}
