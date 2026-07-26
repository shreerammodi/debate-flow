# AGENTS.md

Guidance for AI agents working in this repository.

## Project

**ebb** is a local-first, privacy-centric, keyboard-first app for flowing
competitive debate rounds. It is a flow _editor_: each flow is a `.ebb` file on
the user's own filesystem, reached through the `FlowFs` port in
`src/lib/persistence/` and the narrow Rust commands in
`src-tauri/src/flowfile.rs`. There is no backend and no database. The app is
built as a static export, which Tauri consumes as its frontend; that export is
not deployed anywhere, so a browser is a development target rather than a
product surface, and `flowFsMemory` exists to serve it.

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
    invited. Shared editing (`docs/specs/2026-07-26-shared-editing.md`) is the
    only such feature. It sits behind a master switch that is off by default,
    like `cardmirrorEnabled`, and off leaves every route dead.
  - **The opt-in is an invariant, so it is test-proven, not asserted.** With the
    switch off, the app binds no endpoint, dials no peer, publishes no
    discovery record, and contacts no relay. A test asserts each of those four
    against a fake transport. DNS-based peer discovery stays disabled in every
    state, so an idle ebb publishes nothing about itself.
  - A peer link carries one round and nothing else: no folder listing, no path
    access, no arbitrary read. Hold it to the standard
    `docs/security-review.md` sets for the loopback bridge.
- **All flow I/O goes through the `FlowFs` port** (`src/lib/persistence/flowFs.ts`),
  never directly through `invoke` or a Tauri plugin. That is what lets the
  session, recents, and migration be tested against `flowFsMemory` instead of a
  mocked IPC layer. `tauri-plugin-fs` is deliberately not installed: flow I/O
  uses the five narrow commands in `src-tauri/src/flowfile.rs` so the webview
  never holds a general filesystem capability.
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
- Prefer `git rebase` over `git merge` when integrating changes to maintain a
  linear history.
