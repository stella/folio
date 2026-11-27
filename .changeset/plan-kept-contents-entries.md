---
"@stll/folio-core": patch
---

`applyTemplatePlan` keeps a template's table of contents without the entries it caches: the field stays in the paragraph it begins in, closed behind its separator, and the further paragraphs of its entries are not placed. The entries named the headings and pages of the template's sample text.
