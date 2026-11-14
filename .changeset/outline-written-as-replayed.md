---
"@stll/folio-core": patch
"@stll/docx-core": patch
---

A shape outline written from the model, such as one read from a VML text box, is written as an empty element when it has no child elements, so saving an unedited document a second time gives the same bytes as the first. A line end whose `type` the source omits keeps no type in the model and is saved without one.
