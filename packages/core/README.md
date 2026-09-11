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
| `@stll/folio-core`          | the headless public API — document creation, representation-neutral comparison, the document model, AI-suggestion primitives, and ProseMirror plugins              |
| `@stll/folio-core/markdown` | DOCX ↔ Markdown conversion                                                                                                                                         |
| `@stll/folio-core/server`   | DOM-free document review, explicit tracked edits, comparison, creation, and package helpers                                                                        |
| `@stll/folio-core/*`        | the source-mirrored module tree (e.g. `@stll/folio-core/types/document`, `@stll/folio-core/prosemirror/schema`) for adapters that need lower-level building blocks |

## New documents and reusable style sets

Create a legal-shaped blank document with the built-in preset:

```ts
import { createDocx, createEmptyDocument, createStellaStyleDocumentPreset } from "@stll/folio-core";

const document = createEmptyDocument({
  preset: createStellaStyleDocumentPreset(),
});
const docx = await createDocx(document);
```

Style sets are content-free JSON values. Inspect a source file before presenting
styles for selection, extract the selected dependency closure, persist the
result, and load it into any later document:

```ts
import {
  createEmptyDocument,
  extractDocumentStyleSetFromDocx,
  inspectDocumentStylesFromDocx,
} from "@stll/folio-core";

const catalog = await inspectDocumentStylesFromDocx(sourceDocx);
const styleSet = await extractDocumentStyleSetFromDocx(sourceDocx, {
  name: "Firm contract styles",
  styleIds: catalog.styles
    .filter(({ role }) => role === "default" || role === "quick")
    .map(({ styleId }) => styleId),
});

const document = createEmptyDocument({ styleSet });
```

Extraction excludes document content, metadata, relationships, media, comments,
and revision data. It keeps only the selected styles and the numbering, theme,
font-table, and settings data required to reproduce their formatting.

## Representation-neutral comparison

`compareContent` compares ordered blocks from any document model without first
serializing them to another format. Map durable source anchors to stable block
IDs and include structural ancestry when blocks live in containers:

```ts
import { compareContent, type FolioContentInputBlock } from "@stll/folio-core";

type SourceBlock = {
  anchorId: string;
  type: "clause" | "heading";
  text: string;
  sectionId: string;
};

const toFolioBlocks = (blocks: readonly SourceBlock[]) =>
  blocks.map(
    (source) =>
      ({
        identity: { type: "authoritative", id: source.anchorId },
        kind: source.type,
        text: source.text,
        containerPath: [
          {
            kind: "section",
            identity: { type: "authoritative", id: source.sectionId },
          },
        ],
      }) satisfies FolioContentInputBlock<SourceBlock["type"]>,
  );

const result = compareContent({
  base: { blocks: toFolioBlocks(baseRevision) },
  revised: { blocks: toFolioBlocks(revisedRevision) },
});
if (result.isErr()) throw result.error;

for (const event of result.value.events) renderComparisonEvent(event);
```

Events contain an owned canonical projection of the declared input fields and
are already in full-document render order. Additional caller metadata is not
enumerated; use the returned stable identity to look it up in the source model.
Modified and edited-move segments use UTF-16 offsets compatible with JavaScript
string slicing; move halves share one `move` object, and every table row or
column event references its shared `change`. `FOLIO_CONTENT_COMPARISON_LIMITS`
publishes the block, text, attribute, container, run, and result ceilings. Input
ceilings are checked before alignment; the result ceiling returns the same typed
`FolioContentComparisonLimitError` while constructing the ordered stream.

## Native DOCX redlines

Generate a reviewable `.docx` whose text and supported inline-formatting
differences are native tracked changes:

```ts
import { compareDocx } from "@stll/folio-core";

const result = await compareDocx(originalDocx, revisedDocx, {
  author: "Reviewer",
  timestamp: "2024-03-01T00:00:00.000Z",
});
if (result.isErr()) throw result.error;

await store(result.value.buffer);
console.log(result.value.changes, result.value.verification, result.value.unsupported);
```

For deterministic operations against one document, use `FolioDocxReviewer`
from `@stll/folio-core/server`. Its operations default to tracked changes and
can be enumerated, accepted, or rejected before saving. `getChanges()` includes
inline edits, formatting, paragraph marks, and paragraph, section, table, row,
and cell property changes.

## License

[Apache-2.0](./LICENSE)

folio began as a private fork of [Eigenpal](https://eigenpal.com)'s
[docx-editor](https://github.com/eigenpal/docx-editor); the original license and
copyright are preserved in [`NOTICE.md`](./NOTICE.md).
