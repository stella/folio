---
"@stll/folio-core": patch
---

Add horizontal cell borders on top of a table row's `atLeast`/`auto` minimum height instead of absorbing them. When an explicit row height (ECMA-376 §17.4.81 `w:trHeight` without an `hRule`, or `hRule="atLeast"`) exceeds the cell content, the border thickness now extends the row as Word renders it, fixing cumulative vertical drift in tables of short, bordered rows.
