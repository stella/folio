---
"@stll/folio-core": patch
---

Read through an inline wrapper in find-and-replace and in the compatibility
inspector. Both walked paragraph content with an `if`-chain that ended in a
silent default, so every member neither named was skipped: a phrase inside a
`w:bdo`, a `w:dir`, a smart tag, a run-level `w:customXml` or a `w:ins` could be
read on the page and not found, and an opaque drawing inside one of them was
invisible to the inspector, which then reported a document safe to edit that was
not. Both walks are now a `switch` with a `never` default, so a content type
added to the model has to be given a decision. The search projection counts
every text node the editor's own searchable text counts, deleted text included:
the offsets it produces are resolved back to an editor position, so a member
counted on one side and not the other shifts every later offset and the
replacement lands on the wrong characters.
