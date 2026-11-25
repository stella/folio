---
"@stll/folio-core": patch
"@stll/docx-core": patch
---

A canonical merge drops its separator only where, as the merge runs, every paragraph after the joined one is pending removal: a deletion the batch asks for but refuses, or makes only after the merge, no longer takes the separator away, so accepting a tracked or suggested merge reads as applying it directly. `@stll/docx-core/ops` exports `paragraphHoldsContent` and `paragraphHoldsContentOnceAccepted`.
