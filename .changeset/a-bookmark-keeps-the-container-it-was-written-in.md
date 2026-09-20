---
"@stll/docx-core": minor
"@stll/folio-core": minor
---

Keep a bookmark marker in the container it was written in, so its range still covers what the author selected.

`CT_Body`, `CT_Tc`, `CT_SdtContentBlock`, `CT_Row` and `CT_Tbl` each declare `w:bookmarkStart` and `w:bookmarkEnd` beside their own children, and Word writes them there whenever the selection was whole blocks, whole cells, whole rows or a whole table. folio re-anchored every one of them into a neighbouring paragraph. The element still reached the saved part, so nothing looked lost; what changed was the extent. A bookmark spanning a row came back inside one cell's paragraph, and a `REF` field or a link resolving it then covered the wrong text.

`BlockContent` gains `BookmarkStart` and `BookmarkEnd` as members, so a marker on a body, a cell or a block content control is a block in its own right: it sits between the same two siblings in the model, in the ProseMirror document and in the saved part, with no index to keep honest. `TableRow` and `Table` gain `bookmarks`, a marker plus its position among the cells or rows, because neither models a child a marker could be. These stay typed rather than joining the verbatim sink: folio pairs a start with its end over the model, and a half kept as bytes leaves the other half unpaired and deleted on the first edit — which is the commoner shape, since a bookmark that opens on a row usually closes inside a cell.

The editor gains a block-level `blockBookmarkBoundary` node, the block twin of the inline `bookmarkBoundary` atom, and a row's and a table's markers ride their node's attributes by reference the way an attribute remainder does. The boundary integrity pass reads all four carriers, so a pair spanning two levels stays whole.
