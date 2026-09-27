# Synthetic parity fixtures

## Edit expectation seeds

`build-edit-expectation-fixtures.ts` creates six small synthetic DOCX inputs:

- `edit-final-paragraph-seed.docx`: a target paragraph at the document end;
- `edit-paragraph-boundary-seed.docx`: two adjacent paragraphs with distinct
  text;
- `edit-numbering-seed.docx`: a multilevel list with a nested item;
- `edit-merged-table-seed.docx`: a table with a vertical merge;
- `edit-comment-range-seed.docx`: a comment anchored to a range;
- `edit-notes-fields-sections-seed.docx`: footnote and endnote references,
  complex and simple fields, a section transition, and heading styles.

The seeds use fixed ZIP timestamps and invented text. Rebuild or check them with:

```sh
bun run parity:build-edit-expectation-fixtures
bun run parity:build-edit-expectation-fixtures --check
```

`edit-operation-scripts.ts` defines one saved edit per seed. The focused
`editOperationScripts.test.ts` test runs every script and reads its saved
structure: blocks and labels, table cells, revisions, comment anchors, note
text, and package carrier counts. Reviewed structural expectations for the
final-paragraph, paragraph-boundary, numbering, merged-table, and
notes/fields/sections cases are pinned by `editStructuralExpectations.test.ts`.
The comment-reply script retains a reproducible saved output while its
expectation is being reviewed.

Inspect a saved result and its structural view in the ignored `.cache` directory:

```sh
bun run parity:run-edit-case insert-continuing-numbering
```

## Layout corpus

The layout corpus provides three deterministic, generated fixtures:

- `isolated-page-furniture.docx` isolates full-width wrapped header and footer
  artwork.
- `pairwise-layout-interactions.docx` is generated from a deterministic
  strength-two covering array. Its cases combine section modes, anchor frames,
  wrapping, pagination controls, table modes, and typography.
- `layout-kitchen-sink.docx` adds tables, merged cells, fields, footnotes,
  numbering, tabs, bidirectional text, columns, page borders, and mixed run
  formatting.

All text, links, identifiers, and numeric values are synthetic. The generator
uses fixed ZIP timestamps, so committed fixtures are byte-for-byte stable.
`layout-interaction-matrix.json` records every generated case and its
content-addressed ID; the fixture prints that ID on the corresponding case page.

The current matrix covers every valid pair from 6,500 valid Cartesian
scenarios in 32 generated cases. Inline anchors pair only with inline wrapping;
floating anchors pair only with floating wrap modes.

```sh
bun run parity:build-layout-corpus
bun run parity:check-layout-corpus
bun run parity:layout-matrix
```

The matrix command generates one temporary DOCX per case. Isolating the cases
makes failures attributable to a single axis combination and prevents a section
transition in one case from changing a later case. The temporary documents and
local reports contain synthetic values only.

The isolated and pairwise fixtures should remain small enough to attribute a
regression. Add a feature to the kitchen sink only after it has focused
coverage; the combined file is a discovery tool, not a substitute for a
minimal regression.

## Line-endpoint fixtures

`word-hyphenation-hanging.docx` is generated entirely from the hand-written
OOXML in `build-word-line-endpoint-fixtures.ts`. It contains synthetic text
only. The fixed ZIP timestamps make repeated builds byte-for-byte stable so
the paired manifest can validate the exact DOCX SHA-256.

`word-hyphenation-hanging.word-lines.json` was captured from the Word and
`mutool` versions recorded in the manifest. It covers:

- US and British English automatic hyphenation;
- Czech automatic hyphenation;
- paragraph-level hyphenation suppression;
- the all-caps and consecutive-line hyphenation controls;
- Japanese kinsoku and explicit hanging-punctuation controls;
- adjacent closing punctuation and inline formatting-run boundaries;
- a document-specific prohibited line-start replacement list;
- automatic hyphenation across an inline formatting-run boundary.
- common justification, indentation, tab-stop, numbering, table-cell, and mixed-format layouts.

Slovak remains covered by deterministic dictionary unit tests, not this Word
baseline. Word hyphenation for a language depends on the proofing dictionaries
installed with the local Office installation, and the capture environment did
not provide Slovak automatic hyphenation.

See `../README.md` for the rebuild, capture, and validation commands. A Word
manifest is reviewed interoperability evidence, not an OOXML conformance
oracle.
