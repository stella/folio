---
"@stll/folio-core": minor
---

A template plan can give a part of the document a page other than the template's, such as a landscape annex.

- A source item takes `endsSection`: its paragraph is the last of a section. `TemplatePlan.lastSection` describes the section the body ends with.
- Each such section is built from the template's last section: the same headers, footers and other properties, with only the plan's page (orientation, size, optionally margins) and start (`nextPage`, `evenPage`, `oddPage`) put in place. The first-page header and a page-number start stay with the document's first section.
- Verification holds the output to the plan's sections and their pages (`section-mismatch`); `report.counts.sectionEnds` counts them. A table, a nested paragraph and a block joined into the paragraph before it cannot end a section, and the template needs a last section to build from.
