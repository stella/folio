<p align="center">
  <img src="https://raw.githubusercontent.com/stella/folio/main/.github/assets/banner.png" alt="folio" width="100%" />
</p>

# @stll/folio-react

A Word-document editor for the browser. It opens a real `.docx`, lets you edit
it, and writes a real `.docx` back — preserving pagination, tables, headers and
footers, tracked changes, and footnotes.

The React editor is one layer on top of
[`@stll/folio-core`](https://www.npmjs.com/package/@stll/folio-core), the
headless OOXML parser, document model, and page-layout engine.

Part of [stella](https://github.com/stella/stella), an open-source legal workspace.

## Install

```sh
bun add @stll/folio-react react react-dom use-intl
```

`@stll/folio-core` is installed automatically as a dependency.

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

## Exports

| Import                         | What it is                                                                                                                                              |
| ------------------------------ | ------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `@stll/folio-react`            | the React editor and its components                                                                                                                     |
| `@stll/folio-react/editor.css` | the single bundled stylesheet — import once; it `@import`s the document fonts from `@fontsource/*` (installed as deps), so no external network requests |

For headless, framework-neutral use (OOXML parsing, the document model, the
layout engine, DOCX ↔ Markdown), depend on `@stll/folio-core` directly.

## Peer dependencies

`react` ^18 or ^19 · `react-dom` ^18 or ^19 · `use-intl` >=4

## Acknowledgements

folio began as a private fork of [Eigenpal](https://eigenpal.com)'s
[docx-editor](https://github.com/eigenpal/docx-editor) by
[Jedr Blaszyk](https://github.com/jedrazb). The code has since been extended
(mostly to match the needs of stella). After the upstream repository was taken
down, we're publishing folio in case the fork is useful to others as well. The
original license and copyright are preserved in [`NOTICE.md`](./NOTICE.md).

## License

[Apache-2.0](./LICENSE)
