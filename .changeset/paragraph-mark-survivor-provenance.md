---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

A tracked insertion that records which paragraph rejecting its paragraph mark keeps (a block inserted after the paragraph ending its block list) now saves that record as its own ignorable review-provenance attribute, `folio:rejectionSurvivor`, and reads it back. Rejecting the insertion after a save and reopen now keeps the same paragraph, under the same id, as rejecting it before the save. A reader that does not know the attribute drops only the marker and joins the paragraphs as for any rejected mark.
