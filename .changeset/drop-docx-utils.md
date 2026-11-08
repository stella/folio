---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

`@stll/docx-core` and `@stll/folio-core` no longer depend on `@stll/docx-utils`. The OOXML namespace table they read from it, `OOXML_NS` (with its `OoxmlPrefix` type), now comes from `@stll/docx-core/xml/namespaces`, with the same entries. Installing folio no longer brings in `@stll/docx-utils` or `jszip`.
