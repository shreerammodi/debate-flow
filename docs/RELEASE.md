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

| Product         | What it is                             | Distribution                      | How it updates                               |
| --------------- | -------------------------------------- | --------------------------------- | -------------------------------------------- |
| **Web build**   | Static export in `out/`                | Vercel CDN, no backend            | User reloads; the CDN always serves current  |
| **Desktop app** | Tauri 2 shell wrapping the same `out/` | Signed installers, GitHub Release | In-app signed auto-updater, tournament-gated |

Local-first invariant holds for all three: no backend, no telemetry. The only
network the runtime touches is fetching `latest.json` + update artifacts from
GitHub Releases, and that path is **desktop-only** and **opt-in** (`isDesktop()`
short-circuits the whole update layer on web).

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

- **`web`** — `npm ci -> npm test -> npm run lint -> npm run build`. Gates all
  logic, including the pure update-policy tests in `src/lib/update/*.test.ts`.
- **`desktop`** — `npm ci -> npm run build -> cargo check`. The web build runs
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
the signed artifact in the background, then applies it **only when safe**:

- Background checks are **opt-in** (off by default) and run on launch + every 6h
  once enabled. "Check now" in Settings always works.
- A staged update surfaces as a subtle "Update ready - Restart" chip
  (`UpdateChip`); clicking relaunches into the new version. It never nags
  mid-round.
- Applying is held during the weekly **blackout window** (default Fri->Mon) and
  whenever **Tournament Mode** is on. Downloading is never gated - only the
  restart-to-apply is.
- A release marked `critical: true` can bypass the hold, but only through an
  explicit confirm modal (`CriticalUpdateModal`), never silently.

Full state machine, eligibility layers, and failure properties:
[`desktop/ci-and-deployment.md`](desktop/ci-and-deployment.md) sections 7-10.

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

For a critical fix, edit the published release's `latest.json` to set
`"critical": true` and upload it again over the existing asset (`tauri-action`
does not set it). Clients read the flag on their next check.
