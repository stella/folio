---
"@stll/folio-core": minor
"@stll/folio-agents": patch
---

`insertAfterBlock` / `insertBeforeBlock` accept `outlineLevel` (a direct `w:outlineLvl`, or `null` to clear the anchor's). `compareDocx` states it on every inserted or relocated paragraph, so a body paragraph placed beside a heading that states its own outline level no longer accepts as a heading, and its self-check compares each block's resolved kind and heading level. The agent operation schemas advertise `outlineLevel` for insertions and `setBlockParagraphProperties`.
