---
"@stll/folio-core": minor
---

Preserve explicit OOXML page-break run and field ownership across editor conversion, revision resolution, and DOCX round trips. Retain namespaces on nested run-property revisions and refuse page-break layouts that cannot be projected losslessly. AI edit snapshot anchors now require a structural-boundary fingerprint so edits cannot cross page-break topology changes.
