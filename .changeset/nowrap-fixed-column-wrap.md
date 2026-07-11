---
"@stll/folio-core": patch
---

Wrap `w:noWrap` table cells whose content overflows a pinned column. When the
table layout is fixed or an explicit `w:tblW` (`dxa`/`pct`) width pins the
columns, Word cannot honor `w:noWrap` by widening the column, so it wraps the
content. Measurement now measures such cells at their real column width (instead
of an unbounded width), and the painter drops `white-space: nowrap` for them so
the painted height matches. Auto-width tables still keep `w:noWrap` cells on a
single line. This corrects under-measured rows that previously let extra rows
fit per page and dropped a page.
