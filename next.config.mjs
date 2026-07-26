import { readFileSync } from "fs";
import path from "path";
import { fileURLToPath } from "url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

// The desktop build reads its version from the Tauri runtime, which does not
// exist in a browser. Injecting it here gives `npm run dev` the real number
// instead of the 0.0.0 placeholder getCurrentVersion falls back to.
const { version } = JSON.parse(readFileSync(path.join(__dirname, "package.json"), "utf8"));

/** @type {import('next').NextConfig} */
const nextConfig = {
    output: "export",
    // Emit `route/index.html` (not `route.html`) so Tauri's asset server can
    // resolve a hard load / reload of `/flow`. Without it that route 404s on
    // any non-client navigation and WKWebView shows its native "This page
    // couldn't load" error (e.g. after an updater relaunch).
    trailingSlash: true,
    outputFileTracingRoot: __dirname,
    env: {
        NEXT_PUBLIC_EBB_VERSION: version,
    },
    images: {
        unoptimized: true,
    },
};

export default nextConfig;
