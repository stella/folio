---
"@stll/docx-core": minor
"@stll/folio-core": patch
---

A document whose review provenance folio cannot read now opens instead of being refused. Review provenance is the `folio:` attributes in the ignorable review-history namespace, and the cases are another version, an unknown field, or a malformed value. The reader drops that provenance, reads the revisions it annotates as they are, and reports each dropped value as a `review-provenance-dropped` parse warning. What folio writes is still validated strictly.
