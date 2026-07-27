## Repository Specifics

folio is a Bun-first TypeScript monorepo for browser-based Word-document (`.docx`)
editing. Its published packages have explicit ownership boundaries:

- `@stll/docx-core` (`packages/docx-core`) owns the typed OOXML document model,
  validation, serialization, legal-source compiler, and portable DOCX projection
  kernel. The kernel's Rust source lives in `crates/docx-kernel`; the package exposes
  its browser binding through `@stll/docx-core/projection`.
- `@stll/folio-core` (`packages/core`) owns DOCX parsing, ProseMirror integration,
  framework-neutral editor behavior, and page layout.
- `@stll/folio-react` and `@stll/folio-vue` (`packages/react`, `packages/vue`) are thin
  framework adapters over folio-core.
- `@stll/folio-nuxt` (`packages/nuxt`) provides the SSR-safe Nuxt integration.
- `@stll/folio-agents` (`packages/agents`) provides agent tooling over the public
  editor contracts.
- `packages/playground` and `packages/playground-vue` are private test applications;
  they are not published.

### Commands

- `bun install`
- `bun run build`
- `bun run typecheck`
- `bun run test`
- `bun run test:property`
- `bun run lint`
- `bun run format:check`
- `bun run validate-dist`
- `bun run test:interactions`
- `bun run test:e2e:vue`
- `bun run test:e2e:parity`
- `bun run test:differential`

### Working Rules

- **Preserve upstream attribution.** folio is a fork of the Eigenpal docx-editor (see
  `NOTICE.md`). `NOTICE.md`, `LICENSE`, and the `eigenpal` / `docx-editor` attribution
  comments must stay verbatim; never scrub them.
- Keep React and Vue public contracts in parity. Run
  `bun run check:parity-contract` and `bun run check:export-parity` when an adapter
  contract changes.
- Return minimal data from public APIs; do not export types with no consumer.
- Keep DOCX projection semantics in `crates/docx-kernel`. TypeScript bindings may
  initialize WebAssembly, preserve its versioned result, and translate boundary
  errors; they must not contain a fallback OOXML parser or duplicate Rust logic.
- Keep the portable DOCX kernel single-threaded and browser-native: no WASI, workers,
  shared memory, `SharedArrayBuffer`, or cross-origin-isolation requirement. Run
  `bun run check:docx-kernel` after Rust changes; the gate checks deterministic
  generation, runtime constraints, and transfer-size budgets.
- Resolve OOXML elements by namespace URI plus local name, never by a hard-coded
  prefix. Support Strict and Transitional namespace profiles explicitly, bound ZIP
  and XML resource use, and preserve package paragraph identifiers as facts rather
  than treating them as durable application or host-navigation identities.
- **Add a changeset for any published-package `src` change.** Select every affected
  package, choose the appropriate bump, and add a one-line summary. Use
  `bunx changeset --empty` only when the source change intentionally needs no release.
  The private playground packages need no changeset.
- **Never delete or regenerate `bun.lock` to apply package version bumps.** Run
  `bun scripts/check-lockfile-workspace-versions.ts --write`, then
  `bun install --frozen-lockfile`. The synchronizer owns cached workspace
  self-versions; dependency-graph changes belong in an explicit install.

## Cursor Cloud specific instructions

Toolchain and standard commands live under `## Repository Specifics` → `### Commands`;
this section records only non-obvious cloud-VM caveats.

- **Toolchain lives in the user profile, not the base image.** Bun `1.3.14` is at
  `~/.bun/bin`; Node is provided by nvm (default `v22.22.2`). A login shell sources
  `~/.bashrc` and activates both. A plain non-login `bash -c` falls back to
  `/exec-daemon/node` (Node 22.14).
- **Build and dist validation need Node ≥ 22.18.** Use the nvm default Node before
  running commands that load tsdown configuration.
- Playwright Chromium is optional locally and required for interaction, visual, and
  adapter e2e suites. `python-docx` is optional for differential tests, which skip
  when it is unavailable.
- `.ai/shared` is a submodule. Initialize it before running `bun run sync-ai`; do not
  fetch a floating shared revision during normal CI.
