---
"@stll/docx-core": patch
---

Deleting a table column, directly or as a suggestion, takes the bookmarks, content controls and review records wholly inside the removed cells with them, and a column edit on a table with a row already suggested for insertion or deletion keeps that row's suggestion while the row gains or loses its cell. A bookmark, comment or move reaching past the column still stops the deletion.
