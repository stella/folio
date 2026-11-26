---
"@stll/folio-core": minor
---

A block manifest says where its document starts a page or ends a section, and a template plan can start a paragraph on a new page.

- `DocxBlock` gains `pageBreaks` (page breaks among a paragraph's runs, also in a paragraph without text), `pageBreakBefore` (the paragraph's own formatting or its style's) and `endsSection`.
- `DocxBlockManifest` gains `sections`: each section's `start` (next page, even or odd page, next column, continuous), orientation, page size, margins and column count. A template inspection's sections carry the same facts beside their headers and footers.
- A source item of a plan takes `pageBreakBefore`. The paragraph is written with `w:pageBreakBefore`, verification holds the output to it (`page-start-mismatch`), and the report counts such paragraphs in `counts.pageStarts`. A table and a block joined into the paragraph before it cannot take one.

Direct paragraph formatting and section breaks of a source are not carried over as before; a plan can now keep the page starts they gave.
