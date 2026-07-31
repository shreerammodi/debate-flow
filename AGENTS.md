# AGENTS.md

Guidance for AI agents working in this repository.

## Project

**ebb** is a local-first, privacy-centric, keyboard-first app for flowing
competitive debate rounds. It is a flow _editor_: each flow is a `.ebb` file on
the user's own filesystem, reached through the `FlowFs` port in
`src/lib/persistence/` and the narrow Rust commands in
`src-tauri/src/flowfile.rs`. There is no backend and no database. The app is
built as a static export, which Tauri consumes as its frontend; the same export
is also served as a website, so it is a product surface and not only a
development target. A browser cannot reach the filesystem commands, so the web
build runs on `flowFsMemory`, which keeps flows in memory and mirrors them to
`localStorage`. That makes the browser a place a debater can lose work, not a
place a round lives. Its security headers, including the CSP a static export
cannot get from Tauri, are in `vercel.json`; the desktop shell's stricter
policy is in `src-tauri/tauri.conf.json`.

Run `npm test` and `npm run lint` before considering a change complete.
Formatting is `oxfmt` (via `npm run format` / `format:check`), not Prettier.

## Conventions

- **Tests live under `test/`**, mirroring the `src/` tree: the suite for
  `src/<path>/X.ts(x)` is `test/<path>/X.test.ts(x)` and imports it via the
  `@/` alias. Most `lib/` modules have a test; keep new logic covered and
  test-driven where practical.
- **Pure logic goes in `src/lib`**, not in components. Keep `lib/`
  framework-agnostic and testable; components wire it to React.
- **Local-first**: never add a network call the user has not explicitly opted
  into. The rule is consent, not a ban. Storage is unconditional: every flow is
  a file on the user's own disk, and the user can read, write, and open one with
  no network at all.
    - **Never, under any toggle**: telemetry, analytics, crash reporting, remote
      config, update pings outside the signed updater, accounts, or a backend of
      our own.
    - **Permitted, behind an explicit opt-in**: sending a flow to a peer the user
      invited. Shared editing is the only such feature, specced in
      `docs/superpowers/specs/2026-07-26-shared-editing-design.md`. It sits behind
      a master switch, `collabEnabled`, that is off by default, and off leaves
      every route dead.
    - **On is not the same as reachable, and a launch is not consent.** The
      master switch unlocks Share and Join; it binds nothing on its own, so a
      cold launch with shared editing on says nothing to anyone. Staying bound
      between rounds so a saved contact's invite can land is the one route that
      reaches the network with no round in hand, so it is its own switch,
      `collabListenEnabled`, off by default. Turning it on is what mints the
      macOS local network prompt and the Windows firewall prompt, which is the
      point: the prompt should arrive at the moment a debater asks to be
      reachable, never during startup. Settings shows "Your ID" from the
      `collab_endpoint_id` command, which reads the public half of
      `identity.key` off the disk - never bind an endpoint to learn an id.
    - **Shared editing is desktop only, and that is not a gate to relax.** A
      session is an iroh endpoint, which a browser cannot bind. There is no web
      adapter for `PeerLink` and there should never be one: a stand-in that
      satisfied the port would mint tickets nobody can redeem and tell a
      debater they are connected to a peer that cannot exist. `collabLive()`
      answers "is this offered here", build and switch together, and every
      route that can start a session asks it; `createPeerLinkFor` throws off
      the desktop as the backstop. `collabSettings()` is the switches alone,
      for code that has already been handed a transport, which is what lets the
      suite drive the whole protocol against `peerLinkMemory`.
    - **The opt-in is an invariant, so it is test-proven, not asserted.** With the
      switch off, the app binds no endpoint, dials no peer, publishes no
      discovery record, and contacts no relay. `test/lib/collab/optIn.test.ts`
      runs one session request against a recording transport with the switch on
      and again with it off: on, the recorder shows a bound endpoint carrying an
      explicit discovery and relay config plus a dial per known peer; off, it is
      empty, so the positive control is what makes the empty recorder mean the
      gate held. A relay is only reachable through a transport, and off there is
      none. Two gates rather than one, because they answer different
      questions: `collabLive()` is what the routes a debater takes by hand
      ask, `startForRound` (`src/lib/collab/runtime.ts:166`) and `joinRound`
      (`join.ts:78`), and `collabSettings()` is what code already holding a
      transport reads, at `inviteListener.ts:65`, `persist.ts:35`, and
      `runtime.ts:264`. Every one of them is held to the off case beside its
      own behavior in `join.test.ts`, `inviteListener.test.ts`,
      `runtimeInvites.test.ts`, and `persist.test.ts`; the two that also hold
      the idle listener to the case where shared editing is on and
      `collabListenEnabled` is off are `inviteListener.test.ts` and
      `runtimeInvites.test.ts`. Reopening a round that was shared before is
      the one route no debater asks for, since a `.ebb` carries its peers in
      its sidecar and a double-click is the whole gesture, so `resumeSession`
      reads `collabSettings().listen` before it reaches `startForRound`'s
      `collabLive()`: shared editing on and Listen for invites off binds
      nothing on a cold launch. That route is held to both off cases in
      `runtimeInvites.test.ts`, under the positive control the other four
      have. Switching the master off while a session is already running tears
      it down rather than waiting for the next route to ask
      (`useInviteWatch.ts:42-47`, `test/lib/collab/useInviteWatch.test.tsx`),
      and one thrown while a session is still binding is caught on the far side
      of the await, so a bind that completes after the switch went off ends
      itself instead of coming up with the switch off.
      DNS-based peer discovery stays disabled in every state, so an idle ebb
      publishes nothing about itself.
    - A peer link carries one round and nothing else: no folder listing, no path
      access, no arbitrary read. Hold it to the standard
      `docs/security-review.md` sets for the loopback bridge.
    - **An EndpointId names a peer; it does not route to one.** With DNS
      discovery off, the only lookup left is mDNS, which answers across a room
      and no further, so a dial that carries an id alone spends its whole
      deadline waiting for an address that never comes and then fails. What
      makes a round work between two networks is a relay URL travelling by
      hand: a ticket carries the host's, and a peer already met carries its own
      back on the connection, where it is kept beside the round in its sidecar
      and beside the peer in the contact table. That is addressing, not a
      registry - nothing is published, and an idle ebb still says nothing about
      itself anywhere. Never close this gap by registering a discovery service.
    - **A collab command that waits on the network must be `async`.** Tauri
      runs a `#[tauri::command]` declared without it on the main thread, and a
      bind, a dial and a stop all wait on the network for seconds - which on
      the main thread is a frozen window. `collab.rs`'s `off_thread` puts the
      work on a blocking thread, which is also what makes driving the
      collaboration runtime with `block_on` legal: tokio panics doing that from
      a thread running async tasks. Both halves are held by the suite.
- **All flow I/O goes through the `FlowFs` port** (`src/lib/persistence/flowFs.ts`),
  never directly through `invoke` or a Tauri plugin. That is what lets the
  session, recents, and migration be tested against `flowFsMemory` instead of a
  mocked IPC layer. `tauri-plugin-fs` is deliberately not installed: flow I/O
  uses the six narrow commands in `src-tauri/src/flowfile.rs` so the webview
  never holds a general filesystem capability. `src-tauri/src/sidecar.rs` adds
  the two that persist a collaboration replica, and neither of those takes a
  path either.
- Keyboard-first UX is a core product value - preserve and extend keybindings
  rather than replacing them with mouse-only flows.

## Comments

Distilled from a 2026-07 audit that pruned drifted comments across `src/`.
A comment earns its place by stating a why, an invariant, or a non-obvious
edge case; otherwise leave the code bare.

- **Describe current behavior, in present tense.** Never narrate the change
  you are making: no "now uses X", "no longer stored", "previously",
  "replaces the old X". Those go in the commit message; the comment states
  the surviving fact ("deleting an argument never vaporizes the answers
  written under it"), not its history.
- **No plan or ticket artifacts.** Task numbers, milestone labels, and spec
  references ("(Task 6)", "M3:", "per spec") mean nothing to a future reader.
  This applies to section banners and test `describe` titles, not just prose.
- **No conversation echoes or thinking-out-loud.** Nothing addressed to a
  reviewer ("as requested"), and no left-in self-corrections ("... - wait,
  actually it calls orphanNode"). Resolve the thought, write the conclusion.
- **Update docs when the code they describe changes.** JSDoc drifts silently
  during refactors: after the `order` field became `row`, three navigation
  docs still said "order"; a test kept "Re-import to pick up the new
  navigator.platform" after the re-import was removed. When renaming a field
  or restructuring, grep comments for the old name too.
- **Attach JSDoc to the symbol it documents**, immediately above it - not
  above a neighboring cache variable or a sibling declaration.
- **No mirror comments** that restate the adjacent line ("// clear on
  unmount" on a cleanup assignment), and no hedges that misstate behavior
  ("assumes the group exists" on code that safely no-ops).
- **Section banners are a deliberate convention** (`// --- Actions ---...`);
  keep them, and keep directive comments (`eslint-disable`,
  `@ts-expect-error`) intact.
- **Use the model's vocabulary** (`src/lib/model/types.ts` is the source of
  truth): a **sheet** is a page of the flow (not "page"/"tab"); a **node** is
  an `ArgumentNode` datum; a **cell** is the grid slot at `(speechId, row)` -
  never conflate node and cell; a **speech** is a column identity ("column"
  is fine for the visual dimension); **side** is aff/neg while **role** also
  includes judge. In export code, disambiguate the app's `Sheet` from Excel
  worksheets explicitly.
- **A side is always aff/neg in the model, never in the UI.** The `Side` type
  is aff/neg for every event, but what the user reads comes from
  `sideLabels(event)`: Parliamentary calls the same two sides Gov and Opp and
  its debater slots PM/MG and LO/MO, not 1A/2A and 1N/2N. Never hard-code
  "Aff"/"Neg" into a surface that has the round in hand - the buttons, the
  round header, the ballot, and the exported workbook all read the event's
  labels. Settings is app-wide with no round to ask, so it stays aff/neg.
- **Use plain text**. Never add symbols, glyphs, or other unicode characters.
  Only use standard ASCII characters, unless absolutely necessary. When
  representing keyboard modifiers, use standard terms Meta, Alt, Ctrl, Shift
  instead of glyphs.

## Notes for agents

- There is no server; `npm run build` produces a static site in `./out`.
- The `?` dialog (`src/components/palette/KeybindingsCheatsheet.tsx`) is a
  keyboard-shortcut reference only; conceptual/workflow docs live on the
  external site at https://ebb.smodi.net (outside this repo).
- When adding a UI primitive, follow the existing `components/ui` (shadcn-style)
  patterns; `components.json` configures the generator.
- **Binding a printable key (no Ctrl/Meta) to a command is a trap.** With the
  grid focused and no editor open, Handsontable "fast edits" the selected cell
  on any printable keydown, opening an empty editor that commits over the cell
  and erases it before the app command even runs. `HotGrid`'s `beforeKeyDown`
  guard resolves such chords (bare keys, Alt+key) against the keymap and runs
  them itself so the grid never touches the cell; keep new printable bindings
  flowing through that path, not around it.
- **`Meta+O` is Insert Cell, not Open.** Both `Meta+o` and `Meta+O` belong to
  the cell and row insert commands a debater uses mid-speech, so `flow.open`
  carries no editor chord and no menu accelerator; the start screen binds a
  bare `o` instead. Check `presets.ts` and `reserved.ts` before claiming a
  chord - flowing owns most of the letter space.
- **`Meta+N` is New Window, not New Flow.** `flow.new` carries no editor
  chord and no menu accelerator, for the same reason `flow.open` does not -
  the start screen binds a bare `n` instead, and the command palette or the
  File menu still reach it by click.
- **Every window is a fully independent app instance; there is no "main"
  window.** `src-tauri/src/windows.rs` builds every window at runtime
  (`tauri.conf.json`'s `app.windows` is deliberately empty) with a unique
  `win-N` label; opening a `.ebb` from the file manager always creates a new
  window rather than steering an existing one. `windows::target_window`
  resolves "the window the user is looking at" for the handful of things
  that must reach exactly one window - a native menu action, the CardMirror
  bridge, a shared-editing session - since Tauri's own `emit` broadcasts to
  every open webview by default. `shutdown.rs` generalizes the flush-before-
  exit handshake the same way: closing one window flushes and closes only
  it; quitting (Cmd+Q, or closing the last window) flushes every open window
  and exits only once all of them confirm.
- Prefer `git rebase` over `git merge` when integrating changes to maintain a
  linear history.
