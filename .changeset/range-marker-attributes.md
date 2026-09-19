---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

Keep every attribute a range marker arrived with. `w:moveFromRangeStart` and `w:moveToRangeStart` lost `w:author` and `w:date`, which their schema type requires, so a saved document was markup Word repaired; `w:displacedByCustomXml` was lost on every bookmark, comment range and move range. The markers now model the schema's own `CT_MarkupRange` / `CT_Bookmark` / `CT_MoveBookmark` chain and share one reader and one writer.
