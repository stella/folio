---
"@stll/folio-core": patch
---

Emit schema-valid OOXML for compared packages, and date them from the comparison.

Four shapes a compared package could carry were not shapes the schema allows,
and the round-trip self-check could not see any of them: it reads which words a
view resolves to, and all four parse back to the right words.

- **A revision `w:id` is unique across the package.** One logical change
  serializes as several physical wrappers — a word-level redline cut around the
  words that survived — and the pass that gives each its own id ran only on the
  full repack. The selective save, which rewrites the changed paragraphs and
  re-emits every other part verbatim, exited past it, so a paragraph-local edit
  shipped several revisions under one id. Both exits now run the pass, and every
  id it lets stand or mints goes through one choke point that refuses a
  duplicate.
- **`w:tbl` is `w:tblPr`, `w:tblGrid`, then rows.** Both are required and both
  precede every row; they were emitted only when the model had something to put
  in them, so a table the comparison creates — no authored properties, no
  measured column widths — opened with its first `w:tr`. They are now always
  written, with a grid column per column the widest row spans.
- **`w:hyperlink` wraps a revision, not the other way round.** Run-level
  `w:ins`/`w:del` take run-level content, which a hyperlink is not. Deleting or
  inserting linked text now emits
  `<w:hyperlink><w:del><w:r><w:delText>…`, splitting the revision at each link
  boundary, and reading such a package back restores the same model.
- **A paragraph id is 31-bit.** `w14:paraId`, `w14:textId` and the comment-part
  ids that reference a paragraph are `ST_LongHexNumber` with a maximum below
  `0x80000000`. Producers exist that ignore the bound, and folio preserves the
  ids a document arrives with, so an out-of-range id travelled straight through
  a save. One mapping, a pure function of the id, now brings such a value into
  range — applied when a paragraph is parsed and again across every part of the
  package on the way out, so a paragraph and every reference to it move
  together and a document's identity does not shift between reading and
  writing.

`compareDocx` also restamps `dcterms:modified` in `docProps/core.xml` from its
`timestamp` option. The save wrote the wall clock there, so two runs over
identical inputs with identical options differed in that part alone — the last
clock in a call whose whole contract is that it has none.
