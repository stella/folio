---
"@stll/folio-core": patch
---

Serialize captured XML directly instead of through a parallel node tree. `elementToXml` built a second copy of every subtree in fast-xml-parser's builder format before writing it, and captures nest, so the same bytes were copied at every level on the way out. It now appends into one shared buffer. The output is unchanged: a differential property test compares it against the builder it replaced over generated trees, and the two agree on every one of the 194,413 elements in a 120-package corpus sample.
