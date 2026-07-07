<p align="center">
  <img src="https://raw.githubusercontent.com/stella/folio/main/.github/assets/banner.png" alt="folio" width="100%" />
</p>

# @stll/folio-core

The headless, framework-neutral core of [folio](https://github.com/stella/folio):
the OOXML (`.docx`) parser, the document model, the ProseMirror integration, and
the page-layout engine. It opens a real `.docx`, models it, paginates it, and
writes a real `.docx` back — preserving tables, headers and footers, tracked
changes, and footnotes.

There is **no React** in the import graph, so the core runs on a server or under
any framework. The React editor lives in
[`@stll/folio-react`](https://www.npmjs.com/package/@stll/folio-react), and the
Vue editor lives in [`@stll/folio-vue`](https://www.npmjs.com/package/@stll/folio-vue).

Part of [stella](https://github.com/stella/stella), an open-source legal workspace.

## Install

```sh
bun add @stll/folio-core
```

## Exports

| Import                      | What it is                                                                                                                                                         |
| --------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `@stll/folio-core`          | the headless public API — `createEmptyDocument`, `createDocx`, the document model, AI-suggestion primitives, ProseMirror plugins                                   |
| `@stll/folio-core/markdown` | DOCX ↔ Markdown conversion                                                                                                                                         |
| `@stll/folio-core/server`   | DOM-free helpers for server-side use (block-id derivation, document creation, re-zip)                                                                              |
| `@stll/folio-core/*`        | the source-mirrored module tree (e.g. `@stll/folio-core/types/document`, `@stll/folio-core/prosemirror/schema`) for adapters that need lower-level building blocks |

## License

[Apache-2.0](./LICENSE)

folio began as a private fork of [Eigenpal](https://eigenpal.com)'s
[docx-editor](https://github.com/eigenpal/docx-editor); the original license and
copyright are preserved in [`NOTICE.md`](./NOTICE.md).
