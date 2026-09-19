---
"@stll/folio-core": patch
---

Represent a `w:pict` once. A legacy VML group carrying a text box was claimed by two owners: the run parser kept the whole `w:pict` as one raw drawing, and the text-box pass rebuilt its first `v:textbox` as an editable shape beside it. A save wrote both, so the box's text appeared twice in the saved document, and twice again on every later save. The text-box pass now asks the run parser's own predicate whether a pict is already claimed instead of re-deriving the answer from the markup.
