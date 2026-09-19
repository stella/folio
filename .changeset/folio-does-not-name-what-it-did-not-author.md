---
"@stll/docx-core": minor
"@stll/folio-core": minor
---

Stop minting `wp:docPr@name`. A shape, text box or picture whose name the model
did not carry was written back as `Shape 3`, `TextBox 3` or `Picture 3`, so a
connector named `直接箭头连接符 2` came back in English through the editor round
trip and no later reader could tell a generated name from an authored one. The
serializer now writes only what the model holds, and the insert command names
the object it creates; a drawing with no name writes `@name=""`, the required
attribute with nothing in it.

`wp:docPr@descr` (alt text) and `@title` were never modelled for shapes and text
boxes at all, so a rebuild dropped them: `Shape` and `TextBox` gain `alt` and
`title`, the ProseMirror shape and text-box nodes carry them along with the
authored name, and one reader/writer pair owns all three attributes for every
drawing kind.
