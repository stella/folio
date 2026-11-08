---
"@stll/folio-core": patch
"@stll/folio-react": patch
"@stll/folio-vue": patch
---

A read-only editor no longer changes comment threads from the review sidebar in the ProseMirror session: resolve, reopen, reply and delete report a read-only edit attempt and publish nothing, as the canonical session already did. Template preview reads only the values the host set itself, so a marker named after an inherited property such as `{{constructor}}`, or a value that is neither text nor text spans, keeps its authored text.
