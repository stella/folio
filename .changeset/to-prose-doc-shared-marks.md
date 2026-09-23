---
"@stll/folio-core": patch
---

Speed up opening documents: `toProseDoc` builds each distinct run mark once per conversion, and validation reuses per-mark and per-node attr checks.
