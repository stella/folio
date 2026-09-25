---
"@stll/folio-core": patch
---

Keep the runs of a complex field that a paragraph does not close in source order on save. A TOC opened in its first entry's paragraph no longer saves with its `begin`, instruction and `separate` after the entry's `w:hyperlink`, and an outer field whose result holds a nested field keeps its `begin`, instruction and `separate` instead of losing them.
