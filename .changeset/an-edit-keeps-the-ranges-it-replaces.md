---
"@stll/folio-core": patch
---

Keep the annotations a replaced span carried. ProseMirror drops a mark declared
`inclusive: false` once a replacement reaches the end of what the mark covers,
which is right for formatting and wrong for a mark that names something outside
itself. Replacing a block's text therefore dropped its `comment` marks — taking
the comment's only range start with them, so the save wrote a `commentRangeEnd`
with no start, invalid OOXML and a comment anchored to nothing — and dropped
its `hyperlink` mark, leaving prose that had been a link pointing nowhere.

A replacement now carries the comments and the link its span held (in tracked
mode too, so accepting the change keeps them), keeps the zero-width anchors
inside it (comment references, bookmark boundaries, text-box anchors) rather
than deleting them, and a selective save refuses to patch a story whose comment
ranges it would leave with one half, falling back to a full repack instead.
Every non-inclusive mark now has a recorded disposition, so a new one cannot
join the schema without a decision.
