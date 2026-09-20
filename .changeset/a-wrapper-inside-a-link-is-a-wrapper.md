---
"@stll/docx-core": minor
"@stll/folio-core": minor
---

Read a transparent wrapper inside a link or a simple field as the wrapper it
is. `CT_Hyperlink` and `CT_SimpleField` are both `EG_PContent`, which declares
`w:bdo`, `w:dir`, `w:smartTag` and the run-level `w:customXml`, so
`w:hyperlink > w:bdo > w:r` is markup a producer may write; folio captured all
four whole, which kept the markup and made every run inside it opaque bytes the
editor could not touch. `Hyperlink["children"]` and `SimpleField["content"]`
now carry `InlineWrapper`, the wrapper's children are walked by the container's
own handler map — so a `w:ins` inside a `w:bdo` inside a link is captured for
the same reason a `w:ins` directly inside the link is — and the runs reach the
editor carrying both the link mark and the `inlineWrapper` stack. Saving from
the editor writes the canonical order the wrapper design fixed, revision then
wrapper then hyperlink then run, so a link authored inside a `w:bdo` comes back
with the `w:bdo` around it; the link, the wrapper and the text all survive, and
a paragraph nobody edited keeps its authored order because selective save
replays its bytes. A simple field keeps the wrapper inside itself, because its
node holds its own inline content.
