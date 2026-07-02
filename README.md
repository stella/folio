<p align="center">
  <img src=".github/assets/banner.png" alt="@stll/folio" width="100%" />
</p>

# folio

A Word-document editor for the browser. It opens a `.docx`, lets you edit
it, and writes a `.docx` back — preserving pagination, tables, headers and
footers, tracked changes, and footnotes.

The OOXML parser, document model, and page-layout engine are React-free, so they
run on a server or under any framework. The React editor is one layer on top.

Part of [stella](https://github.com/stella/stella), an open-source legal workspace.

## Packages

This is a [bun](https://bun.sh) workspace monorepo with two published packages:

| Package                                 | What it is                                                                                                                                                                                        |
| --------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| [`@stll/folio-core`](./packages/core)   | the headless, framework-neutral core — OOXML parsing, the document model, the ProseMirror integration, and the layout engine; no React in the import graph, so non-React adapters can build on it |
| [`@stll/folio-react`](./packages/react) | the React editor and its components, built on `@stll/folio-core`                                                                                                                                  |

## Install

```sh
# the React editor (pulls in @stll/folio-core)
bun add @stll/folio-react react react-dom use-intl

# or just the headless engine
bun add @stll/folio-core
```

## Usage

```tsx
import { DocxEditor } from "@stll/folio-react";
import "@stll/folio-react/editor.css";

export function Editor({ docx }: { docx: ArrayBuffer }) {
  return <DocxEditor documentBuffer={docx} onSave={(out) => download(out)} />;
}
```

The editor renders to the DOM; under SSR, load it from a client-only/dynamic
import.

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

Bundled locales: `en`, `de`, `fr`, `es`, `cs`, `ar`, `et`, `hu`, `lt`, `lv`,
`pl`, `pt-BR`, `sk`. Arabic (`ar`) is right-to-left: set `dir="rtl"` on a
container around the editor for that locale.

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
bun run build       # builds both packages (core first)
bun run typecheck
bun run test        # unit suite for both packages
bun run lint
bun run validate-dist   # clean-room publish-shape validation for both packages
```

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
