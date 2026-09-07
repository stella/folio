---
"@stll/folio-core": patch
---

Emit schema-valid OOXML for compared packages, and date them from the comparison.

Three shapes a compared package could carry were not shapes the content model
allows, and the round-trip self-check could not see any of them: it reads which
words a view resolves to, and all three parse back to the right words.

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

`compareDocx` also restamps `dcterms:modified` in `docProps/core.xml` from its
`timestamp` option. The save wrote the wall clock there, so two runs over
identical inputs with identical options differed in that part alone — the last
clock in a call whose whole contract is that it has none.
