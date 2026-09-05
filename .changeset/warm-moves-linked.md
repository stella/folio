---
"@stll/folio-core": minor
"@stll/folio-agents": minor
---

A relocation is written into the document as a linked pair.
`deleteBlock`, `insertAfterBlock` and `insertBeforeBlock` take a `moveId`; when
one id names exactly one deletion and one insertion the applier writes
`w:moveFrom` and `w:moveTo` instead of an unrelated deletion and insertion.
An id that does not is reported as an `unpairedMove` normalization and both
halves apply plainly.

`compareDocx` emits the pair, so a reordered document now says so to every
OOXML consumer rather than only in its JSON change list. A relocated paragraph
is recognized when it keeps at least 80% of its word tokens, so a clause edited
on the way to its new home is still a move.

`FolioAIEditNormalization` is a discriminated union on `code` rather than one
shape with a `splitMultilineText`-specific field.
`FolioDocumentOperationResult.nextRevisionId` is optional: a host bridge that
delegates to an editor it does not control omits it rather than guessing.
