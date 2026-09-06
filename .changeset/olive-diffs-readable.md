---
"@stll/folio-core": minor
---

`diffWordSegments` stops shredding a rewritten paragraph. An LCS maximises
matched characters, so a rewritten sentence used to come back as a dozen
struck-through fragments interleaved with a dozen inserted ones. Three rules
pull it back: a match made only of separators is not a match, a one-token match
with changes on both sides of it is dropped into them, and a paragraph whose
surviving matches are too short for its length is replaced whole. On a
320-paragraph rewrite the package carries 68% fewer separately marked runs; a
light edit is marked word by word exactly as before.

The diff now takes options: `granularity` (`"word"` default, or `"character"`
to mark the changed letters inside a token) and `normalization` (`case`,
`whitespace`). `granularity` is threaded through `applyFolioDocumentOperations`
and `compareDocx`; normalization is not, because a comparison that leaves a
difference unmarked does not accept back to the target.
