---
"@stll/folio-core": patch
---

`docxToMarkdown` keeps the number of a numbered heading: a `Heading 2` numbered through its style or its own `w:numPr` renders `## 1. Scope` instead of `## Scope`, counted with the list items around it, so the Markdown shows the same number `getContent()` labels the block with. A heading whose numbering level hides its marker (`w:vanish`) still renders `## Scope`, and a numbered heading no longer joins the list that follows it.
