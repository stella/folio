---
"@stll/folio-core": patch
---

A page break before an empty section-break paragraph now starts a new page before a continuous section. Empty section-break paragraphs that are numbered, bordered, or the only paragraph of their section keep their line, and one that follows its section's content ignores `w:pageBreakBefore`.
