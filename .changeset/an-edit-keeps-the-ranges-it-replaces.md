---
"@stll/folio-core": patch
---

Keep the annotations a replaced span carried. Replacing a block's text dropped
every `comment` mark on it, because the mark is `inclusive: false` and the
replacement reached the end of what it covered: the comment's only range start
went with it, and the save wrote a `commentRangeEnd` with no start, which is
invalid OOXML and a comment anchored to nothing. A replacement now carries the
comments its span held, keeps the zero-width anchors inside it (comment
references, bookmark boundaries, text-box anchors) rather than deleting them,
and a selective save refuses to patch a story whose comment ranges it would
leave with one half, falling back to a full repack instead.
