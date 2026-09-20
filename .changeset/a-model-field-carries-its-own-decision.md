---
"@stll/docx-core": patch
---

Record the reserved-value decision for every field the model declares.

`ImageWrap.distanceSlots`, `ImageWrap.polygon` and the font table's two verbatim sinks landed without an entry in the registry their types are total over, so `typecheck:reserved-values` did not compile. Each carries markup rather than a value with a reserved meaning, and now says so.
