---
"@stll/folio-core": minor
---

With `sourceHeaders: "keep"` the headers and footers of the source's later sections are no longer lost. An output section after the first that starts with a source section showing other content gets new parts of its own: the template's content for those pages followed by that section's, named by that output section alone. A later section the output merges into another is counted in `leftOut.laterSectionParts`, and each kept part of a later section names its output `section`.
