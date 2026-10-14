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
bun add @stll/folio-react react@^19.2.0 react-dom@^19.2.0 use-intl@^4.0.0 \
  prosemirror-commands@^1.7.1 prosemirror-dropcursor@^1.8.2 prosemirror-gapcursor@^1.4.1 prosemirror-history@^1.5.0 \
  prosemirror-inputrules@^1.5.0 prosemirror-keymap@^1.2.3 prosemirror-model@^1.25.9 prosemirror-state@^1.4.4 \
  prosemirror-tables@^1.8.5 prosemirror-transform@^1.12.0 prosemirror-view@^1.42.6 y-prosemirror@^1.3.7 \
  y-protocols@^1.0.6 yjs@^13.6.31
```

ProseMirror and Yjs packages are required peers, including the collaboration
packages `y-prosemirror` and `y-protocols`. Install one compatible copy of each in the host
application so schemas, plugins, editor states, and collaborative documents share
the same runtime. The command above includes the complete editor peer closure.

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

`react` ^19.2 · `react-dom` ^19.2 · `use-intl` >=4, plus the ProseMirror
and Yjs peers listed in the installation command above.

Keyboard shortcuts default to the editor root. Set `keyboardShortcuts="document"`
only when this editor should own shortcuts across the host page.

## Acknowledgements

folio began as a private fork of [Eigenpal](https://eigenpal.com)'s
[docx-editor](https://github.com/eigenpal/docx-editor) by
[Jedr Blaszyk](https://github.com/jedrazb). The code has since been extended
(mostly to match the needs of stella). After the upstream repository was taken
down, we're publishing folio in case the fork is useful to others as well. The
original license and copyright are preserved in [`NOTICE.md`](./NOTICE.md).

## License

[Apache-2.0](./LICENSE)

See [document search](../../docs/document-search.md) for highlighting, navigation, and result callbacks.
