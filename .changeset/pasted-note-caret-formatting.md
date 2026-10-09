---
"@stll/folio-core": minor
"@stll/folio-vue": patch
---

Breaking: `fromProseDoc` now requires an explicit stylesheet source; pass `{ stylesheetSource: { type: "package" } }` or `{ stylesheetSource: { type: "supplied", styles } }` with the stylesheet used for projection. Preserve pasted formatting and authored overrides through save.
