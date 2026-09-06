---
"@stll/folio-core": minor
---

A page says where a click can land, not just what it paints.

`DisplayPage` gains `regions`: a tree of boxes over the same primitives — the
content area, header and footer slots, notes, paragraphs, lines, empty lines,
tables, rows and cells — each carrying the model range and the story it
resolves to, plus the block id, comment threads and row and column indices a
surface reads. The producer opens a region around the painting that fills it, so
the structure comes from the walk that lays the page out rather than from a
second pass over it, and a region indexes into the paint list rather than
copying it.

The DOM backend paints each region as the element the interaction layer has
always looked for, with the runs nested inside, so clicks, drags and selections
resolve against what the producer laid out. A glyph run also states when it is a
line-edge space run the line was fitted without, and what one of those spaces
would have advanced, which is what lets a caret step through them.

`interactionContract.test.ts` reads the interaction layer's own source, extracts
every class and data attribute it looks for, and fails unless each one has a
source in the IR. A reader that learns a new selector without the producer
gaining something to emit it from fails that test.
