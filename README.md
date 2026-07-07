<p align="center">
  <img src=".github/assets/banner.png" alt="@stll/folio" width="100%" />
</p>

# folio

A framework-neutral DOCX engine and browser editor toolkit. It opens a `.docx`,
lets you edit it, and writes a `.docx` back — preserving pagination, tables,
headers and footers, tracked changes, and footnotes.

The OOXML parser, document model, and page-layout engine are framework-neutral,
so they run on a server or under any UI adapter. The React and Vue editors are
thin layers on top, with a Nuxt module for Vue/Nuxt apps and an agent-tooling
package for document review workflows.

Part of [stella](https://github.com/stella/stella), an open-source legal workspace.

## Packages

This is a [bun](https://bun.sh) workspace monorepo with these published packages:

| Package                                 | What it is                                                                                                                                                                                             |
| --------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| [`@stll/folio-core`](./packages/core)   | the headless, framework-neutral core — OOXML parsing, the document model, the ProseMirror integration, and the layout engine; no framework UI in the import graph                                     |
| [`@stll/folio-react`](./packages/react) | the React editor and its components, built on `@stll/folio-core`                                                                                                                                       |
| [`@stll/folio-vue`](./packages/vue)     | the Vue 3 editor and composables, built on the same core and layout engine                                                                                                                             |
| [`@stll/folio-nuxt`](./packages/nuxt)   | a Nuxt 3/4 module that registers the Vue editor safely for Nuxt apps                                                                                                                                   |
| [`@stll/folio-agents`](./packages/agents) | framework-neutral LLM tool definitions and executors for reading `.docx` documents and proposing tracked changes or comments through `@stll/folio-core`                                             |

## Install

```sh
# the React editor (pulls in @stll/folio-core)
bun add @stll/folio-react react react-dom use-intl

# the Vue editor
bun add @stll/folio-vue vue

# Nuxt integration
bun add @stll/folio-nuxt

# agent/review tooling
bun add @stll/folio-agents

# or just the headless engine
bun add @stll/folio-core
```

## Usage

```tsx
import { DocxEditor } from "@stll/folio-react";
import "@stll/folio-react/standalone.css";

export function Editor({ docx }: { docx: ArrayBuffer }) {
  return <DocxEditor documentBuffer={docx} onSave={(out) => download(out)} />;
}
```

The editor renders to the DOM; under SSR, load it from a client-only/dynamic
import.

### Styling: two integration modes

The editor's chrome (toolbar, menus, dialogs) is authored with Tailwind utility
classes and semantic design tokens. Pick the stylesheet that matches your app.

**1. `standalone.css` — no Tailwind required (default).** A single,
self-contained import: the bundled document/ProseMirror styles plus a
pre-compiled copy of every utility folio's own components use. The utilities are
scoped under `.folio-root`, so they cannot leak into or restyle your app, and
the design tokens ship as low-specificity fallbacks you can override.

```tsx
import "@stll/folio-react/standalone.css";
```

Override the theme by setting the tokens on `.folio-root` (a normal class rule
outranks the shipped `:where(.folio-root)` fallbacks):

```css
.folio-root {
  --background: #fdfdfc;
  --foreground: #1c1c1a;
  --primary: #3b5bdb;
  /* ...only the tokens you want to change... */
}
```

For dark mode, add a `.dark` class to an ancestor (e.g. `<html>`); the editor
and its body-portalled overlays theme themselves from it.

**2. `editor.css` — you already run Tailwind.** Import only the bundled document
+ chrome styles and let your own Tailwind build generate the utilities. Point a
`@source` at folio's shipped code so the classes its components use are scanned,
and supply the semantic tokens (`--background`, `--foreground`, `--popover`, …)
from your design system:

```css
/* your app's Tailwind entry */
@import "tailwindcss";
@source "../node_modules/@stll/folio-react/dist/**/*.js";
```

```tsx
import "@stll/folio-react/editor.css";
```

Use this mode when your app's tokens and folio's should stay in lockstep. Do not
import both stylesheets: `standalone.css` already contains everything
`editor.css` does.

## Internationalization

The editor reads its UI strings from [`use-intl`](https://github.com/amannn/use-intl)
under the `folio.*` namespace, so it must be wrapped in an `IntlProvider`. folio
**bundles its own translations** for that namespace, so a consumer only merges
folio's catalog and sets the locale — the editor localizes itself:

```tsx
import { IntlProvider } from "use-intl";
import { DocxEditor } from "@stll/folio-react";
import { FOLIO_LOCALES, getFolioMessages } from "@stll/folio-react/messages";
import "@stll/folio-react/editor.css";

export function Editor({ docx, locale }: { docx: ArrayBuffer; locale: string }) {
  return (
    <IntlProvider locale={locale} messages={getFolioMessages(locale)}>
      <DocxEditor documentBuffer={docx} />
    </IntlProvider>
  );
}
```

`@stll/folio-react/messages` exports:

- `getFolioMessages(locale: string): FolioMessages` — the bundled `{ folio: … }`
  catalog for `locale`, falling back to English for any locale folio does not ship.
- `FOLIO_LOCALES` — the bundled locales as a readonly tuple (also a `FolioLocale`
  type and an `isFolioLocale` guard).

Bundled locales: `en`, `de`, `fr`, `es`, `cs`, `ar`, `et`, `he`, `hi`, `hu`,
`lt`, `lv`, `pl`, `pt-BR`, `sk`, `tr`, `zh-CN`. Arabic (`ar`) and Hebrew
(`he`) are right-to-left: set `dir="rtl"` on a container around the editor for
those locales.

**Merging with your own app messages.** folio owns exactly one top-level
namespace (`folio.*`), so a shallow spread merges cleanly with your app's other
namespaces. Put folio first so your app wins on any intentional override:

```tsx
const messages = { ...getFolioMessages(locale), ...appMessages[locale] };
```

Do **not** re-declare the `folio.*` keys in your own catalog: let folio's bundled
catalog be the source of truth for that namespace (otherwise a stale app copy
would win the merge). The playground (`packages/playground`) wires this up with a
language switcher you can use to preview every bundled locale.

## Development

```sh
bun install
bun run build       # builds all published packages
bun run typecheck
bun run test        # unit suites across the workspace
bun run lint
bun run validate-dist   # clean-room publish-shape validation for published packages
```

## Releasing

Releases are driven by [Changesets](https://github.com/changesets/changesets).
The published packages are versioned independently.

**Every PR that edits published package source under `packages/{core,react,agents,vue,nuxt}/src`
must add a changeset.** CI enforces this (`bun run changeset:check`):

```sh
bunx changeset          # pick packages + bump level, write a summary
```

For a source change that genuinely needs no release (comments, internal-only
refactor), record that explicitly instead:

```sh
bunx changeset --empty
```

Commit the generated `.changeset/*.md` file with your PR.

How a version reaches npm:

1. PRs merge to `main`, each carrying its changeset(s).
2. `release-pr.yml` maintains a **"Version Packages"** PR that applies the
   pending changesets — bumping the affected `package.json` versions, updating
   changelogs, and re-syncing `bun.lock`.
3. Merging that PR lands the version bumps on `main`, which trips
   `publish.yml`'s package path filter and runs the hardened OIDC publish +
   GitHub Release for the bumped package(s).

Changesets never publishes; `publish.yml` is the sole publish mechanism.

## Acknowledgements

folio began as a private fork of [Eigenpal](https://eigenpal.com)'s
[docx-editor](https://github.com/eigenpal/docx-editor) by
[Jedr Blaszyk](https://github.com/jedrazb). The code has since been extended
(mostly to match the needs of [stella](https://github.com/stella/stella)). After the upstream
repository was taken down, we are publishing the folio fork as an independently maintained
continuation. The original license and copyright are preserved in
[`NOTICE.md`](./NOTICE.md).

## License

[Apache-2.0](./LICENSE)
