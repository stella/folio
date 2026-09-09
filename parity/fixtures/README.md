# Synthetic parity fixtures

## Layout corpus

The layout corpus provides three deterministic, generated fixtures:

- `isolated-page-furniture.docx` isolates full-width wrapped header and footer
  artwork.
- `pairwise-layout-interactions.docx` combines page-relative positioning,
  expanded page furniture, continuous sections, cached page markers, and
  keep-next flow.
- `layout-kitchen-sink.docx` adds tables, merged cells, fields, footnotes,
  numbering, tabs, bidirectional text, columns, page borders, and mixed run
  formatting.

All text, links, identifiers, and numeric values are synthetic. The generator
uses fixed ZIP timestamps, so committed fixtures are byte-for-byte stable.

```sh
bun run parity:build-layout-corpus
bun run parity:check-layout-corpus
```

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
