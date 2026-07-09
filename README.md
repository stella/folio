<p align="center">
  <img src=".github/assets/banner.png" alt="@stll/folio" width="100%" />
</p>

# folio

Browser editor and framework-neutral engine for `.docx` files. It opens, edits,
and writes Word documents while preserving pagination, tables, headers and
footers, tracked changes, and footnotes.

The core package is framework-neutral. React, Vue, Nuxt, and document-review
packages build on top of it.

Part of [stella](https://github.com/stella/stella), an open-source legal workspace.

See [DOCX platform boundary](./docs/docx-platform.md) for what belongs in folio
and how editors, headless tools, agents, and hosts share one document model and
operation contract.

## Packages

This is a [Bun](https://bun.sh) workspace with these published packages:

| Package                                   | Use it for                                                                        |
| ----------------------------------------- | --------------------------------------------------------------------------------- |
| [`@stll/folio-core`](./packages/core)     | OOXML parsing, document model, ProseMirror integration, and page layout           |
| [`@stll/folio-react`](./packages/react)   | A React editor UI built on `@stll/folio-core`                                     |
| [`@stll/folio-vue`](./packages/vue)       | A Vue 3 editor and composables                                                    |
| [`@stll/folio-nuxt`](./packages/nuxt)     | Nuxt 3/4 registration for the Vue editor                                          |
| [`@stll/folio-agents`](./packages/agents) | Document review tools for reading `.docx` files and proposing comments or changes |

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

## Quick Start

```tsx
import { DocxEditor } from "@stll/folio-react";
import "@stll/folio-react/standalone.css";

export function Editor({ docx }: { docx: ArrayBuffer }) {
  return <DocxEditor documentBuffer={docx} onSave={(out) => download(out)} />;
}
```

In SSR apps, load the editor with a client-only or dynamic import.

## Styling

Pick one stylesheet.

Use `standalone.css` when your app does not run Tailwind or you want folio's
styles isolated:

```tsx
import "@stll/folio-react/standalone.css";
```

Override tokens on `.folio-root`:

```css
.folio-root {
  --background: #fdfdfc;
  --foreground: #1c1c1a;
  --primary: #3b5bdb;
  /* ...only the tokens you want to change... */
}
```

For dark mode, add `.dark` to an ancestor such as `<html>`.

Use `editor.css` when your app already runs Tailwind. Add folio's distributed JS
to Tailwind's sources, then import the stylesheet:

```css
/* your app's Tailwind entry */
@import "tailwindcss";
@source "../node_modules/@stll/folio-react/dist/**/*.js";
```

```tsx
import "@stll/folio-react/editor.css";
```

Do not import both stylesheets. `standalone.css` already includes everything in
`editor.css`.

## Internationalization

The editor uses [`use-intl`](https://github.com/amannn/use-intl). Wrap it in an
`IntlProvider` and pass folio's bundled messages:

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

- `getFolioMessages(locale: string): FolioMessages`
- `FOLIO_LOCALES`
- `FolioLocale`
- `isFolioLocale(locale: string): locale is FolioLocale`

Bundled locales: `en`, `de`, `fr`, `es`, `cs`, `ar`, `et`, `he`, `hi`, `hu`,
`lt`, `lv`, `pl`, `pt-BR`, `sk`, `tr`, `zh-CN`. Arabic (`ar`) and Hebrew
(`he`) are right-to-left: set `dir="rtl"` on a container around the editor for
those locales.

To merge folio messages with app messages, keep folio under its own `folio.*`
namespace:

```tsx
const messages = { ...getFolioMessages(locale), ...appMessages[locale] };
```

Do not copy folio's `folio.*` keys into your app catalog.

## Development

```sh
bun install
bun run build
bun run typecheck
bun run test
bun run lint
bun run validate-dist
```

## Releasing

Releases use [Changesets](https://github.com/changesets/changesets). Add a
changeset to every PR that edits published package source under
`packages/{core,react,agents,vue,nuxt}/src`:

```sh
bunx changeset
```

For source changes that do not need a release:

```sh
bunx changeset --empty
```

CI checks this with `bun run changeset:check`. Merging the generated
**Version Packages** PR publishes changed packages through `publish.yml`.

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
