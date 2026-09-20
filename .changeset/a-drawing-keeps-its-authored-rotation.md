---
"@stll/folio-core": minor
---

A drawing's authored rotation and flips survive the editor projection. The
editor carried `a:xfrm`'s three values inside one CSS transform string, which
can state neither `rot="0"` nor `flipH="0"`: both spell the identity, which is
what an absent transform already means, so opening a document and saving it
again dropped the attribute. The `image`, `shape` and `textBox` nodes now carry
each value explicitly (`docxRotation`, `docxFlipH`, `docxFlipV` on `ImageAttrs`,
`ShapeAttrs` and `TextBoxAttrs`; `null` for absent) and write it back, while the
CSS string stays a projection of them for rendering. A rotate or flip from the
editor states every value it decides, so rotating back to zero says zero rather
than handing the decision back to the file.

The attrs are additive: a node persisted without them is read from its CSS
string as before, which is the only record such a node has. The collaboration
attr schema still goes to version 3, because a snapshot a newer build wrote
must not reach an older one that would drop the three keys unread.
