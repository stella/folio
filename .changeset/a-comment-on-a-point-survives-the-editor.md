---
"@stll/folio-core": minor
---

Keep a comment or tracked move anchored on a point, at the point it was anchored at, through the editor.

Word writes a comment left at an insertion point as `w:commentRangeStart` immediately followed by `w:commentRangeEnd`, then the `w:commentReference` run; a tracked move whose range covers nothing is spelled the same way. The editor projects a comment as a `comment` mark over the content its range covers, and a move range as a marker placed around the move's wrappers, so a range with nothing between its two markers had no carrier: the comment still listed, because its reference is an atom of its own, but the anchor position was gone and the save wrote a reference with no range. An empty move range came back around the whole paragraph, which is where the marker carrier puts one whose wrappers it cannot find.

The editor gains a zero-width `rangeAnchor` node holding the two markers, so the pair comes back adjacent at the position it was authored at, inside the wrapper it was authored inside. The pair is one node rather than two boundary nodes: two would leave a position between them for a caret, and typing there would widen a range the reviewer drew as a point. Deleting the anchor removes the comment, exactly as deleting its reference does. A range with any content in it, including one whose content is only a bookmark boundary or a capture, is untouched and still travels as the mark.

The node is additive: editor state written before it holds none and needs no migration.

Sixty container-census pairs — the six range markers across the ten inline containers that declare them — are now `modelled`, thirty-six of them moving from `dropped (editorProjection)` here.
