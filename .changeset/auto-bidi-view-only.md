---
"@stll/folio-core": patch
---

Keep auto-detected paragraph direction out of the saved package. A paragraph whose base direction was only inferred from its text on load (not an explicit direction change, and not a `w:bidi` the source paragraph already carried) no longer writes `w:bidi` on save; the editor still lays the paragraph out right-to-left.
