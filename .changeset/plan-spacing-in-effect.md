---
"@stll/folio-core": minor
---

The block manifest reports the spacing in effect around a paragraph (`spacingInEffect`): the space in front of it and behind it, each side the paragraph's own, else its style's with inheritance resolved, else the document defaults', else none, and its line spacing (`line`, `lineRule`), resolved the same way and single where nothing states one. A side decided by auto spacing or a count of lines is left out. `spacing` still holds only what the paragraph sets by hand. A plan's spacing may now also set the line spacing (`line`, counted by `lineRule`: `auto`, `exact` or `atLeast`; `auto` when left out), which the output writes and verification holds it to. A caller can then keep a table's text as tight and as closely set as its source, where the template's body text would space it out.
