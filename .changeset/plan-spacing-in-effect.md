---
"@stll/folio-core": minor
---

The block manifest reports the space in effect in front of a paragraph and behind it (`spacingInEffect`): each side the paragraph's own, else its style's with inheritance resolved, else the document defaults', else none. `spacing` still holds only what the paragraph sets by hand. A plan can then keep a cell paragraph as tight as its source sets it through its style, where the template would give it the space of its own body text.
