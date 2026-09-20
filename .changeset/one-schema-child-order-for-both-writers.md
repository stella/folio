---
"@stll/docx-core": minor
"@stll/folio-core": patch
---

Write a compiled `w:rPr`'s children in the schema's order, from the same generated list folio-core writes from.

`@stll/docx-core` holds a second `w:rPr` writer — the one the legal-source compiler and the build-from-scratch export share — and it had grown an order of its own, emitting `w:highlight`, `w:sz` and `w:szCs` ahead of `w:rFonts`. A run carrying both a font and a size therefore came out in one order from this package and another from folio-core's serializer. `EG_RPrBase` is an `xsd:choice` referenced `maxOccurs="unbounded"`, so both spellings are valid; what the canonical order buys is one form, the one Word writes, from both writers.

The generated order moves down to where both can read it: `@stll/docx-core/schema` is a new subpath exporting `SEQUENCE_CHILDREN` and the writer that orders by it, and folio-core's declared-child table spreads that same object in. One emitted order, one sort, and a serializer that cannot restate either.
