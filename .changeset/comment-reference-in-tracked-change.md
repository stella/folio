---
"@stll/docx-core": patch
---

Editing a paragraph whose tracked insertion or deletion holds a point comment reference (a `w:commentReference` with no comment range) no longer fails: splitting, inserting into and deleting from that change keep the reference inside it.
