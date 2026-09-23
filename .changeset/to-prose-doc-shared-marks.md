---
"@stll/folio-core": patch
---

Speed up opening documents: `toProseDoc` builds each distinct run mark once per conversion, and validation reuses per-mark and per-node attr checks. Paragraph, table and table-row attr reads are cached and frozen, so `ReadProseMirrorAttrsResult` now returns `issues` as a readonly array.
