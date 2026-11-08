---
"@stll/folio-core": patch
---

The experimental canonical session opens a document that repeats a content control id. Opening the document keeps the first control with each id and gives each later control carrying it an id the package does not use. The same is already done for paragraph ids. Every control and its content is kept. A control that alternate content restates for older readers keeps its id, and a signed package is left unchanged.
