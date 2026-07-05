# Investigation: fork gaps vs. upstream (deferred, on-record)

Compares **our fork** (`@stll/folio-*`, `packages/core` + `packages/react`, plus
in-progress `packages/vue`) against the **upstream reference** (Eigenpal
`@eigenpal/docx-editor-*`, v1.9.0: `core`, `react`, `vue`, `nuxt`, `i18n`, `agents`).

Scope: things our fork does **not** have and is **deferring**. The Vue + Nuxt
adapters are separate in-flight work and are out of scope here. No code changes are
proposed; this is a record of what upstream ships that we don't, and how cheaply each
could be added later.

Upstream paths below are relative to the shallow clone at
`.../scratchpad/upstream/`.

---

## Summary

| Gap                                                                                                                           | Upstream location                                  | Our status                                                                                                           | Recommendation                                                                                                                                                                                                                                  | Rough effort                                                                                                          |
| ----------------------------------------------------------------------------------------------------------------------------- | -------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------- |
| Agents package (`DocxReviewer`, agent tools, MCP surface, `EditorBridge`/`EditorRefLike`, React+Vue chat UI, AI SDK adapters) | `packages/agents` (`@eigenpal/docx-editor-agents`) | Absent. No agent/MCP/chat-UI package. Core has partial pieces: `ai-edits/`, `ai-suggestions/`, `content-controls/`   | Add later as an **independent optional package** (`@stll/folio-agents`). Editors need it only as a type-only `EditorRefLike` contract; the one runtime coupling (React `LocalizedAgentPanel`) is optional. Do **not** block editor parity on it | Large (~46 src files, ~7.7k LOC excl. tests; ~12.7k incl. tests). But additive — near-zero change to existing editors |
| Standalone i18n package with hand-rolled ICU runtime                                                                          | `packages/i18n` (`@eigenpal/docx-editor-i18n`)     | We keep i18n in the React package on `use-intl`; the neutral catalog now lives in `packages/core/src/i18n/messages/` | **Do not fully extract now.** Keep `use-intl` in React; keep the neutral catalog in core (already done for Vue). Revisit only if a non-React/Vue host or per-locale code-splitting becomes a hard requirement                                   | Medium-high if pursued (reimplement + maintain ICU runtime, migrate React off `use-intl`). Zero if deferred           |
| Core `plugin-api` (external plugin host)                                                                                      | `core/src/plugin-api` (`./plugin-api`)             | Absent                                                                                                               | Optional add-on; upstream marks it `@experimental`. Defer until there's a plugin author                                                                                                                                                         | Small (~4 files, ~552 LOC)                                                                                            |
| Core `mcp` server + `docx-editor-mcp` bin                                                                                     | `core/src/mcp` (`./mcp`)                           | Absent                                                                                                               | Optional add-on. Overlaps with the agents-package MCP surface; pick one home if we ever ship MCP                                                                                                                                                | Medium (~4 files, ~1.4k LOC)                                                                                          |
| Core `core-plugins` (registry + docxtemplater plugin)                                                                         | `core/src/core-plugins` (`./core-plugins`)         | Absent; we template via `@stll/template-conditions` instead                                                          | Optional; our templating path differs deliberately. Defer / likely skip                                                                                                                                                                         | Medium (~6 files, ~1.7k LOC)                                                                                          |
| Core `agent` (`DocumentAgent`)                                                                                                | `core/src/agent` (`./agent`)                       | Absent; partial overlap in `ai-edits/`, `ai-suggestions/`, `content-controls/`                                       | Load-bearing **for** the agents package, optional for the editor. Bring in with the agents package if/when we adopt it                                                                                                                          | Large (~15 files, ~5.9k LOC)                                                                                          |
| Core `types/agentApi`                                                                                                         | `core/src/types/agentApi.ts` (`./types/agentApi`)  | Absent                                                                                                               | Type-only companion to `agent`; comes along with it                                                                                                                                                                                             | Small (~630 LOC, types only)                                                                                          |
| Core `editor` spine (`computeLayout`, layout scheduler)                                                                       | `core/src/editor` (`./editor`)                     | **Already covered** by our `controller/` (`folioEditor`, `layoutScheduler`, `layoutPipeline`, `layoutSession`)       | No action; superseded by our own controller                                                                                                                                                                                                     | n/a                                                                                                                   |
| Core `docx/serializer`                                                                                                        | `core/src/docx/serializer` (`./docx/serializer`)   | **Already present** at `packages/core/src/docx/serializer/`                                                          | No action; not a gap                                                                                                                                                                                                                            | n/a                                                                                                                   |

---

## 1. `@eigenpal/docx-editor-agents` (`packages/agents`)

### What it provides

A framework-agnostic agent/AI toolkit plus chat UI, split so backends, React, and Vue
each pull only what they need. Root export (`src/index.ts`):

- **`DocxReviewer`** — headless "Word-like" review API over a raw `.docx` buffer
  (`DocxReviewer.fromBuffer(buffer, author)` → `getContentAsText()`, `addComment()`,
  `replace()` as tracked change, `applyReview({ comments, proposals })` for batching an
  LLM JSON response, `toBuffer()`). Works on a static file with no live editor.
- **Agent tools** — `agentTools`, `executeToolCall`, `getToolSchemas` (OpenAI
  function-calling schema), `getToolDisplayName`; tool defs live in `src/tools/`
  (`breaks.ts`, `formatting.ts`, `pages.ts`, plus `types.ts`).
- **`createReviewerBridge`** — wraps `DocxReviewer` in the `EditorBridge` interface so
  the same tools / MCP server run against a static file.
- **`WordCompatBridge`** — compile-time-only Word JS API parity contract.

### Subpath split (from `package.json` `exports`)

| Subpath                                           | Contents                                                                                                                              | Needs peer       |
| ------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------- | ---------------- |
| `.`                                               | `DocxReviewer`, tools, reviewer bridge, types (no framework)                                                                          | none             |
| `./bridge`                                        | `createEditorBridge`, `EditorBridge`, **`EditorRefLike`**, tools                                                                      | none             |
| `./server`                                        | tools + `DocxReviewer` + `createReviewerBridge` for API routes / serverless / Workers, no React                                       | none             |
| `./react`                                         | `useAgentChat`, `useDocxAgentTools`, `AgentPanel`, `AgentChat*` components                                                            | react            |
| `./vue`                                           | `AgentPanel.vue`, `AgentChatLog/Composer/Timeline/SuggestionChip.vue`, `AIContextMenu.vue`, `AIResponsePreview.vue`, `useAgentBridge` | vue              |
| `./mcp`                                           | `McpServer`, `runStdioServer`, zero-dep JSON-RPC framing (`src/mcp/`)                                                                 | none             |
| `./ai-sdk/server` `./ai-sdk/react` `./ai-sdk/vue` | Vercel AI SDK adapters; shared logic in `src/ai-sdk/shared.ts` (`toAgentMessages`)                                                    | ai / @ai-sdk/vue |

### Peer deps (all optional)

`react`, `vue`, `ai`, `@ai-sdk/vue` are **all** declared `optional: true` in
`peerDependenciesMeta`. Runtime `dependencies` are `docxtemplater`, `jszip`, `pizzip`,
`xml-js` (the DOCX-buffer read/write path for `DocxReviewer`). So the package can be
installed and its `.`/`/server`/`/mcp`/`/bridge` subpaths used with no UI framework at
all.

### Does the Vue/React **editor** need it at runtime?

Both upstream editors declare `@eigenpal/docx-editor-agents: ^1.9.0` as a normal
dependency, but the actual source coupling is almost entirely **type-only**:

- **Vue editor** — type-only. Only `import type { EditorRefLike } from
'@eigenpal/docx-editor-agents/bridge'` in `src/components/DocxEditor/types.ts` and the
  `ref-conformance.test-d.ts` type test. `src/index.ts` just mentions the package in a
  doc comment. No value import anywhere in `packages/vue/src`.
- **React editor** — type-only **except one** value import:
  `src/components/DocxEditor/LocalizedAgentPanel.tsx` does
  `import { AgentPanel } from '@eigenpal/docx-editor-agents/react'`. That is an optional
  convenience wrapper (a localized `AgentPanel`); the ref contract itself is the
  type-only `EditorRefLike` used by `ref-conformance.test-d.ts`.

`EditorRefLike` is a **structural interface** (`getDocument`, `addComment`,
`proposeChange`, `findInDocument`, `getSelectionInfo`, `applyFormatting`,
`insertBreak`, `getPageContent`, `onContentChange`, ...). The bridge also imports
`ParagraphHighlightOptions` / `ScrollToParaIdOptions` **type-only** from
`@eigenpal/docx-editor-core/utils/paragraphFlashTypes` (deliberately from the DOM-free
types module, to keep the agents type surface free of browser code).

### Can folio add this later as an independent optional package?

**Yes.** The editor↔agents coupling is a structural, type-only ref contract, so:

- The editor's imperative ref just needs to satisfy the `EditorRefLike` **shape**; that
  can be asserted with a local `test-d` conformance test that imports the type from a
  future `@stll/folio-agents/bridge`. No runtime dependency on the agents package is
  required in the editor.
- The single runtime coupling (React `LocalizedAgentPanel`) is optional — omit it, or
  let it live in the agents package's `/react` subpath, so the editor package stays
  clean.
- Everything else (`DocxReviewer`, tools, MCP, AI SDK adapters, chat UI) is additive and
  framework-gated behind optional peers.

This fits our fork's "vertical slice" and "keep core framework-neutral" rules: a new
`@stll/folio-agents` package would be a self-contained slice with optional `react`/`vue`
peers, not a change to the editors. The one prerequisite is that our editor exposes an
imperative ref matching an agreed `EditorRefLike`, and that the core exposes the
DOM-free `paragraphFlashTypes`-equivalent types the bridge references.

### Size

46 non-test source files (55 `.ts` + 2 `.tsx` + 7 `.vue` including tests),
~7,722 LOC excluding tests (~12,718 LOC including the `__tests__` suite). Note this
package also **depends on** core's `agent/` subsystem (gap #3 below) at type level, so
adopting the full agents package pulls the `agent` + `types/agentApi` core surfaces
along with it.

---

## 2. Standalone i18n package (`packages/i18n` = `@eigenpal/docx-editor-i18n`)

### What upstream ships

A dependency-light locale package (dev deps only: `tsup`, `typescript`; **no**
`use-intl`, **no** runtime deps). `src/index.ts` (~665 LOC in `src/`) ships a
hand-rolled ICU runtime alongside the locale data:

- **`createT(strings, lang)`** → typed `TFunction` doing dot-path lookup + `{var}`
  interpolation + ICU `{count, plural, =0 {…} one {# …} other {# …}}` via
  `Intl.PluralRules`.
- **`deepMerge(base, override)`** — merges a partial locale over English, treating
  `null` leaves as "untranslated, fall back to English."
- **`formatMessage(template, vars, lang)`** — the interpolation/plural engine (with an
  explicit ReDoS-hardened plural regex).
- **Types**: `LocaleStrings` (derived from `en.json`, the source of truth),
  `PartialLocaleStrings`/`Translations`, `TranslationKey` (every valid dot path),
  `LocaleCode`.
- **Locale JSON at package root** (`en.json` … `zh-CN.json`, `en.json` ~872 lines,
  10 locales) with **per-locale subpath exports** (`./en`, `./de`, `./fr`, `./he`,
  `./hi`, `./id`, `./pl`, `./pt-BR`, `./tr`, `./zh-CN`) for per-locale tree-shaking /
  code-splitting.
- A **codegen block** (`GENERATED START … GENERATED END`, driven by
  `bun run i18n:codegen`) that regenerates the locale imports/registry from the on-disk
  JSON filenames; `i18n:validate` fails CI if the block drifts.

The React and Vue adapters wrap `createT` in framework-native bindings; a non-React/Vue
host calls `createT` directly.

### Our fork today

- i18n lives **inside the React package** on **`use-intl`** (`use-intl >=4.0.0` peer;
  `IntlProvider` in `renderAsync.tsx`, `useTranslations("folio")` in `DocxEditor.tsx`,
  `FormattingBar.tsx`, etc.).
- The **neutral catalog was relocated into core** for the Vue work:
  `packages/core/src/i18n/messages/` (`messages.ts` + `messages.gen.ts` + JSON). folio
  owns exactly one top-level namespace `folio.*`, so a shallow merge into the host's
  messages never collides. `en.json` is ~163 lines (much smaller than upstream's 872 —
  folio's single-namespace surface).
- We ship **17 locales** (ar, cs, de, en, es, et, fr, he, hi, hu, lt, lv, pl, pt-BR,
  sk, tr, zh-CN) vs upstream's 10.

### Trade-off of a full extraction to `@stll/folio-i18n` dropping `use-intl`

**Pros**

- Framework-neutral `t()` usable from core / a CLI / a server renderer / any future
  host with no framework dependency.
- Removes the `use-intl` peer from the React package (one fewer peer to version).
- Per-locale subpath exports enable per-locale code-splitting (import `./de` only).

**Cons**

- We take ownership of an ICU runtime (interpolation, CLDR plurals, `null`-fallback
  merge, ReDoS-safe plural parsing) that `use-intl` maintains for us today — including
  the security-sensitive regex.
- Migrating the React package off `use-intl` (`IntlProvider`, `useTranslations`) to a
  `createT`-based binding touches every translated component.
- Our catalog already lives in core with a clean single-namespace `folio.*` design; a
  separate package would re-solve a problem we've mostly solved for the Vue case.

### Recommendation

**Defer the full extraction.** Keep `use-intl` in React; keep the neutral catalog in
core (already done). The framework-neutral goal is largely met by the core catalog + a
thin per-framework binding. Only extract a standalone `@stll/folio-i18n` with a
hand-rolled runtime if a concrete non-React/Vue host or a hard per-locale
code-splitting requirement lands. If we do, port upstream's ReDoS-hardened
`formatMessage` verbatim rather than writing our own plural parser.

---

## 3. Other upstream-only core surfaces

From upstream `core/package.json` `exports`, cross-checked against
`packages/core/src`. Ordered by how load-bearing they are.

- **`./agent` — `core/src/agent` (`DocumentAgent`, ~15 files, ~5.9k LOC).**
  Headless, framework-agnostic model inspect/edit API (`createAgent`, `executeCommand`,
  `getAgentContext`, selection context, content-control creation, repeating sections).
  **Load-bearing for the agents package**, optional for the editor. We have partial
  overlap in `ai-edits/`, `ai-suggestions/`, `content-controls/`. → Bring in only if we
  adopt the agents package.

- **`./types/agentApi` — `core/src/types/agentApi.ts` (~630 LOC, types only).**
  `Position`/`Range`/command/context types for the agent API. Type-only companion to
  `agent`; comes along with it. Optional otherwise.

- **`./mcp` — `core/src/mcp` (~4 files, ~1.4k LOC) + `docx-editor-mcp` bin.**
  In-core Model Context Protocol server exposing document-editing tools to AI clients
  (`createMcpServer`, `startStdioServer`, `handleJsonRpcRequest`, `core-tools.ts`,
  `cli.ts`). **Optional add-on.** Overlaps with the agents package's own `/mcp` surface
  — if we ever ship MCP we should pick one home, not both.

- **`./core-plugins` — `core/src/core-plugins` (~6 files, ~1.7k LOC).**
  Headless plugin registry (`pluginRegistry`, `CorePlugin`) plus a bundled
  `docxtemplater` plugin and its `mcp-tools.ts`. **Optional**, and our fork deliberately
  templates via `@stll/template-conditions` (`prosemirror/plugins/templateDirectives.ts`)
  rather than docxtemplater, so this is likely a **skip**, not just a defer.

- **`./plugin-api` — `core/src/plugin-api` (~4 files, ~552 LOC).**
  External plugin host contract: `EditorPluginCore`, `PluginPanelProps`, `PanelConfig`,
  `SidebarItem`, `RenderedDomContext`, `resolveItemPositions`. Upstream marks it
  `@experimental` (breaking changes in minors). **Optional add-on**; defer until there's
  a concrete plugin author.

- **`./editor` — `core/src/editor` (~4 files, ~461 LOC).**
  Shared layout `computeLayout` + rAF-coalescing `createLayoutScheduler` +
  `stripScrollFlag`, the "stateful-orchestration spine" for the React/Vue adapters.
  **Already covered** by our `controller/` (`folioEditor.ts`, `layoutScheduler.ts`,
  `layoutPipeline.ts`, `layoutSession.ts`). Not a gap — our controller supersedes it.

- **`./docx/serializer` — `core/src/docx/serializer`.**
  **Already present** in our fork at `packages/core/src/docx/serializer/`. Not a gap.

- **`utils/*` fan-out.** Upstream explodes many `./utils/<name>` subpaths (cardStyles,
  comments, findReplace, fontOptions, headingCollector, highlightColors, listState,
  reportIssue, sidebarConstants, textSelection, units, autoScroll, ...). Our `utils/`
  covers the load-bearing ones (units, headingCollector, replaceText, fonts, colors,
  clipboard, ...). The rest are UI-helper conveniences tied to upstream's React
  toolbar/sidebar; **not needed for editor/parsing parity** and best added on demand
  when a consumer needs the specific helper, rather than mirroring the whole subpath
  fan-out.

**Net:** the only genuinely "missing subsystem" cluster is agent/MCP-related
(`agent`, `types/agentApi`, `mcp`, `core-plugins`, `plugin-api`) — all optional add-ons
that travel with the agents-package decision (gap #1), except `core-plugins` which we've
deliberately diverged from. `editor` and `docx/serializer` are already covered.
