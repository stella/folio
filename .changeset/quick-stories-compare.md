---
"@stll/folio-core": minor
---

`compareDocx` compares every story present on both sides — main, headers,
footers, footnotes and endnotes — instead of the body alone. A pair differing
only in a footnote used to be reported as agreeing.

`FolioDocumentOperationResult` gains `nextRevisionId`: the first revision id a
following batch may allocate against the same document. Word's revision-id
space is the package rather than the part, so a caller writing one batch per
story has to seed each from the previous batch's value; the batch is the only
thing that knows how many ids it took.

`FolioDocxReviewer.acceptAll` and `rejectAll` now sweep every story rather than
the body, so a revision in a header or a note no longer survives an accept-all.

`COMPARE_UNSUPPORTED_REASONS` drops `secondary-story` and gains
`story-not-editable`; a story present on one side only is still reported as
`story-missing-in-base` / `story-missing-in-target`.
