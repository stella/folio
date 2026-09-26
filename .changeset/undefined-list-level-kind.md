---
"@stll/folio-core": patch
"@stll/folio-agents": patch
---

A paragraph numbered at a level its list does not define shows no marker, and every reader now agrees: `getContent()`, the AI snapshot, `getContentAsText()` and the `read_document` rows read it as `kind: "paragraph"` that keeps its `listLevel` and `listReference` (it was a `listItem`, printed with a `•` by `getContentAsText()`), as Markdown already rendered it. It no longer counts toward the list's numbers. A block that is not a heading is a `listItem` exactly when it has a `displayLabel`.
