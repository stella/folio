---
"@stll/folio-core": patch
---

A split's new half reads its properties from the paragraph it was split from without taking that paragraph's source identity, also when the split is redone, so no two paragraphs of a session claim one source paragraph.
