---
"@stll/folio-core": patch
---

A canonical session opens a long document faster and holds less memory: it aligns each paragraph's text with its editor positions when the paragraph is first addressed, not all at open. What a paragraph's editor node alone cannot place, and every paragraph comments run through or mark, is still checked at open; a paragraph whose source and node do not align is refused when it is first addressed.
