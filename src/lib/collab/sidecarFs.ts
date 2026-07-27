/**
 * Where the sidecar is kept.
 *
 * One port, two adapters, in the shape `FlowFs` already uses: Tauri on the
 * desktop and an in-memory map for the browser and the suite. The port takes a
 * round id and never a path, so the webview cannot steer where this writes;
 * the shell resolves the location itself.
 */

import { isDesktop } from "@/lib/update/adapter";

export interface SidecarFs {
    /** Null when this round has no sidecar, which callers treat as ordinary. */
    read(roundId: string): Promise<string | null>;
    write(roundId: string, text: string): Promise<void>;
}

let cached: SidecarFs | null = null;

export async function getSidecarFs(): Promise<SidecarFs> {
    if (cached) return cached;
    // Dynamic on both branches so the browser bundle never pulls in Tauri's JS
    // API, matching how every other desktop touchpoint is gated.
    const mod = isDesktop() ? await import("./sidecarFsTauri") : await import("./sidecarFsMemory");
    cached = mod.createSidecarFs();
    return cached;
}

/** Test seam: swap in a fixture adapter, or reset with null. */
export function setSidecarFs(fs: SidecarFs | null): void {
    cached = fs;
}
