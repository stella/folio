---
"@stll/folio-core": minor
"@stll/folio-agents": minor
---

`setBlockParagraphProperties` changes a block's list level or paragraph style
without touching its words, recorded as a `w:pPrChange` carrying the complete
previous property set so a reject restores it the way Word does. `FolioAIBlock`
gains `listLevel`, and the two insert operations take `listLevel` and a
nullable `styleId` so an inserted paragraph no longer takes its level and style
from whichever block happens to follow it.

`compareDocx` reports the edit as `paragraph-format`. A demoted list item used
to reach the comparison as no change at all, so the redline said two different
documents agreed. Its round-trip self-check now covers each block's style and
list level alongside its text.
