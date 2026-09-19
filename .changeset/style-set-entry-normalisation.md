---
"@stll/folio-core": patch
---

A `DocumentStyleSet` is normalised where it enters folio rather than trusted. A set persisted as JSON before folio learned to repair one could still carry a style numbering it never defines, two styles under a single id, or an initial paragraph style it does not contain, and `createEmptyDocument` would panic on the first of those. It now repairs all three through the owners the parser and the extractor already use, leaves the caller's value untouched, and reports each repair on the resulting document. `DOCUMENT_STYLE_SET_VERSION` is unchanged: the shape did not move.
