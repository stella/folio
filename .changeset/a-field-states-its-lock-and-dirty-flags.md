---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

Keep `@w:fldLock` and `@w:dirty` as a field authored them, on `w:fldSimple` and on the `w:fldChar` that opens a complex field. Three readers and three writers had each collapsed the two attributes to "present and true", so an explicit `w:dirty="0"` -- a field inside a `TOC` result that says not to recompute -- parsed as an absence and saved as one. Both directions now live in one module, `docx/fieldState`. The editor keeps the distinction too: the field node's `fldLock` and `dirty` attrs default to absent rather than `false`, so projecting a field through the editor no longer invents an explicit off. The attributes are written as `1`/`0`, matching Word. Because the persisted attr shape changes, the collaboration attr schema goes to version 2, and `migrateFolioYjsSnapshot` drops the `false` a version-1 snapshot stored for a field that authored neither flag.
