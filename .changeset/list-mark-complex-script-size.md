---
"@stll/folio-core": patch
---

Size a list paragraph's final line from the paragraph mark's `w:szCs` only when the mark is complex script (a right-to-left paragraph, or `w:rtl` / `w:cs` on the mark); left-to-right list paragraphs keep the mark's `w:sz`.
