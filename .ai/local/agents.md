## Repository Specifics

folio is a Bun-first TypeScript monorepo: a Word-document (`.docx`) editor for
the browser, built on ProseMirror. Two published packages plus a dev playground:

- `@stll/folio-core` (`packages/core`) — the headless engine: OOXML parsing, the
  document model, the ProseMirror integration, and the page-layout engine. It is
  **framework-neutral and must stay React-free** (enforced by the
  `react-free-core` + `model-purity` arch tests and the clean-room dist check).
- `@stll/folio-react` (`packages/react`) — the React editor: a thin renderer over
  folio-core. UI logic lives in core managers; the hooks are thin bindings.
- `packages/playground` — a private Vite app that mounts the editor for the
  visual + interaction e2e tests; not published.

### Commands

- `bun install`
- `bun run build` (core first, then react)
- `bun run typecheck`
- `bun run test` (unit suites, both packages)
- `bun run lint`
- `bun run validate-dist` (clean-room publish-shape check)
- `bun run test:interactions` (Playwright behaviour e2e via the playground)
- `bun run test:differential` (folio vs python-docx parse parity; full corpus)
- `bun run format` (oxfmt)

### Working Rules

- **Keep `@stll/folio-core` React-free.** Never import `react`/`react-dom` (or a
  React-package type) into `packages/core`. Framework-agnostic UI logic belongs in
  a core manager (`extends Subscribable`, which documents both a React
  `useSyncExternalStore` and a Vue `watchEffect` binding); the React hook is a
  thin binding only. Declare a minimal structural type in core rather than
  importing one from the React package.
- **Preserve upstream attribution.** folio is a fork of the Eigenpal docx-editor
  (see `NOTICE.md`). `NOTICE.md`, `LICENSE`, and the `eigenpal` / `docx-editor`
  attribution comments must stay verbatim — never scrub them.
- **Fix pre-existing bugs you find** (separate commit) rather than preserving
  known-wrong behaviour to "keep behaviour identical".
- Behaviour is guarded by the interaction e2e suite and the python-docx
  differential parity gate — extend them when you change parsing, layout, or
  editor interactions.
- Return minimal data from public APIs; do not export types that have no consumer.

## Cursor Cloud specific instructions

Toolchain and standard commands live under `## Repository Specifics` → `### Commands`
above; this section only records the non-obvious cloud-VM caveats.

- **Toolchain lives in the user profile, not the base image.** Bun `1.3.14` is at
  `~/.bun/bin` and Node is provided by nvm (default `v22.22.2`). A login/interactive
  shell sources `~/.bashrc` and puts both on `PATH` (and activates the nvm Node). A
  plain non-login `bash -c` does **not**, and falls back to `/exec-daemon/node`
  (Node 22.14). Prefer a login shell, or reference bun explicitly as
  `~/.bun/bin/bun`.
- **`bun run build` / `bun run validate-dist` need Node ≥ 22.18.** tsdown loads
  `packages/*/tsdown.config.ts` with a config loader that requires native TS
  type-stripping. The default `/exec-daemon/node` (22.14) lacks it and the optional
  `unrun` loader is not installed, so building under that node fails with
  `Failed to import module "unrun"`. Run these with the nvm Node 22.22.2 active (it
  is the nvm default; a login shell or `nvm use default` selects it). lint,
  typecheck (`tsgo`), and `bun test` are unaffected by the node version.
- **Dev server:** `bun --filter @stll/playground dev` serves the editor at
  http://localhost:4200 (Vite, `strictPort`). It loads an empty document by default;
  `?file=<name>` loads a fixture from `tests/visual/fixtures` (e.g.
  `?file=sample.docx`). The product is fully client-side — there is no backend,
  database, or external service to start.
- **Optional test dependencies (not installed by the update script):** Playwright
  chromium (`bunx playwright install chromium`) for `test:interactions` /
  `test:visual`; `python-docx` for `test:differential` (the test auto-skips if it is
  missing). `bun run sync-ai` needs the private `.ai/shared` git submodule, which is
  not initialized here, so do not run it.
