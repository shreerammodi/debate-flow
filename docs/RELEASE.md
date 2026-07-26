# Shipping Ebb

The canonical "how Ebb is built, released, and updated" doc. Start here.

Ebb ships **one codebase (`src/`) as two products** from a single version
number. This page is the overview and the release procedure; the deep-dives it
links carry the internals.

- Updater runtime model, policy, failure modes: [`desktop/ci-and-deployment.md`](desktop/ci-and-deployment.md)
- Signing keypair setup + release steps: [`desktop/releasing.md`](desktop/releasing.md)
- First-run trust on unsigned beta builds: [`desktop/manual-trust.md`](desktop/manual-trust.md)

---

## 1. Shipping model

| Product         | What it is                             | Distribution                      | How it updates                                      |
| --------------- | -------------------------------------- | --------------------------------- | --------------------------------------------------- |
| **Web build**   | Static export in `out/`                | Vercel CDN, no backend            | User reloads; the CDN always serves current         |
| **Desktop app** | Tauri 2 shell wrapping the same `out/` | Signed installers, GitHub Release | In-app signed auto-updater, install on user confirm |

Local-first invariant holds for all three: no backend, no telemetry, no
accounts. Two network paths exist, and the user opts into each one:

1. **Updates.** Fetching `latest.json` + update artifacts from GitHub Releases.
   **Desktop-only** and **opt-in** (`isDesktop()` short-circuits the whole
   update layer on web), and the install waits on user confirm.
2. **Shared editing.** A direct peer connection to a partner the user invited,
   behind a master switch that is off by default. Off binds no endpoint and
   contacts nothing, which is asserted by test. See
   `docs/superpowers/specs/2026-07-26-shared-editing-design.md`.

Neither path sends a flow anywhere the user did not choose.

## 2. Versioning

A release is one semver that must move in lockstep across four files, edited by
hand:

- `package.json` -> `version`
- `src-tauri/tauri.conf.json` -> `version`
- `src-tauri/Cargo.toml` -> `version`
- `src-tauri/Cargo.lock` -> refresh with `cargo update -p ebb --manifest-path src-tauri/Cargo.toml`

Then commit all four, tag `vX.Y.Z`, and push with `git push --follow-tags`.
**The pushed tag is what triggers the desktop release** (section 5).

## 3. Continuous integration (`.github/workflows/ci.yml`)

Runs on every push to `main` and every PR. Two parallel jobs, both on
`ubuntu-22.04`:

- **`web`** - `npm ci -> npm test -> npm run lint -> npm run build`. Gates all
  logic, including the pure update-policy tests in `src/lib/update/*.test.ts`.
- **`desktop`** - `npm ci -> npm run build -> cargo check`. The web build runs
  first because Tauri's `generate_context!` reads `frontendDist: ../out`, which
  must exist for `cargo check` to compile.

CI does **not** build full installers (release-only) and does **not** deploy the
web build (Vercel handles that, section 4).

## 4. Web deployment (Vercel)

The web build is a pure static export (`output: "export"`,
`images.unoptimized`) deployed via **Vercel's native Git integration** - not a
GitHub Actions workflow. No deploy tokens or secrets live in this repo.

- Push to `main` -> production deploy.
- Pull request -> isolated preview URL.

Configured once in the Vercel dashboard: framework preset **Next.js**, build
`npm run build`, output dir `out`. Deploying the web build _is_ its update -
there is no update concept on web; a reload is non-destructive (continuous
autosave) and the user controls when they reload.

## 5. Desktop release (`.github/workflows/release.yml`)

**Trigger:** pushing a `v*` tag (section 2 does this for you).

A 3-way build matrix (`fail-fast: false`) - macOS universal, Linux x64, Windows
x64 - runs `tauri-apps/tauri-action`, which builds each installer, signs the
updater artifacts with the Ed25519 key, generates `latest.json`, and uploads
everything to a GitHub Release named `ebb vX.Y.Z`.

The workflow creates the release as a draft, and the `publish` job
(`needs: release`) publishes it after all three platforms succeed. The draft
keeps a partly uploaded release away from clients: the updater reads
`releases/latest/download/latest.json`, and a draft is never "latest". **A tag
push reaches users on its own. If one platform fails, nothing publishes.**

Full procedure and one-time signing-key setup: [`desktop/releasing.md`](desktop/releasing.md).

## 6. Signing (two independent layers)

- **Updater signing (Ed25519) - mandatory, live from day one.** Guarantees
  update integrity independent of the OS. Public key committed in
  `tauri.conf.json`; private key + password are repo secrets
  (`TAURI_SIGNING_PRIVATE_KEY`, `..._PASSWORD`). Lose the private key and you can
  never sign updates again.
- **OS code signing (Apple Developer ID / Windows Authenticode) - deferred but
  pre-wired.** `release.yml` already declares the empty `APPLE_*` secret slots;
  populate them to enable notarized signing with no workflow changes. Until
  then, beta builds are unsigned and users do a one-time trust step
  ([`desktop/manual-trust.md`](desktop/manual-trust.md)).

## 7. How users update

**Web:** reload the page. Whatever the Vercel CDN serves is current. Nothing to
install, no version to track.

**Desktop:** the in-app updater checks GitHub Releases, downloads and verifies
the signed artifact in the background, and installs it **only when the user
says so**:

- Background checks are **opt-in** (off by default) and run on launch + every 6h
  once enabled. "Check now" in Settings always works.
- A staged update surfaces as a subtle "Update ready - Restart" chip
  (`UpdateChip`); clicking installs and relaunches. It never nags mid-round and
  never applies on its own.
- Installing first writes the open flow. If that write fails, the update is
  abandoned and the round stays on screen.

There is no calendar-based gating: an explicit click is the whole safety model,
which is why Tournament Mode and the weekly blackout window were removed in
0.4.0. A user who wants no interruptions at all turns auto-update off in
Settings.

Full state machine and failure properties:
[`desktop/ci-and-deployment.md`](desktop/ci-and-deployment.md).

## 8. Beta ship checklist

Blocking before the first real release:

- [ ] **Replace the placeholder updater pubkey.** The Ed25519 key committed in
      `tauri.conf.json` is a dev placeholder. Generate a production keypair
      (`npm run tauri signer generate`), commit the new pubkey, and set
      `TAURI_SIGNING_PRIVATE_KEY` + `..._PASSWORD` as repo secrets. See
      [`desktop/releasing.md`](desktop/releasing.md).
- [ ] **Connect the repo in the Vercel dashboard** so section 4's Git
      integration is live.

Ships fine for beta, worth doing after:

- [ ] Enable OS code signing (secrets already slotted) to drop the manual-trust
      step for macOS/Windows users.
- [ ] Consider a `pub_date`-based minimum age so a bad publish can't auto-apply
      instantly.

## 9. Cutting a release (quick reference)

```bash
# From a clean main with CI green, set VERSION=X.Y.Z, then by hand:
# - edit `version` in package.json, src-tauri/tauri.conf.json, src-tauri/Cargo.toml
cargo update -p ebb --manifest-path src-tauri/Cargo.toml   # refresh Cargo.lock
git commit -am "$VERSION" && git tag -s "v$VERSION" -m "v$VERSION"
git push --follow-tags
# -> release.yml builds 3 installers into a draft GitHub Release,
#    then publishes it once every platform succeeds. No click needed.
# -> the /latest/ redirect flips; desktop clients pick it up on next check.
```

There is no `critical` flag. Every install waits for the user to click, so a
fix ships as an ordinary release and reaches people the next time they check.
