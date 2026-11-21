---
"@stll/folio-core": patch
---

A package whose `_rels/.rels` is not in the package relationships namespace is refused on open with `MalformedPackageRelationshipsError`, on both sessions; the legacy session now refuses it too.
