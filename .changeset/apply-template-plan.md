---
"@stll/folio-core": minor
---

Add `applyTemplatePlan` to the server entry: it places the blocks of a source document onto a template's own package from a plan of block ids, styles, list levels and reference spans, and returns the output with a report. `analyzeDocxBlocks` and `inspectDocxTemplate` give the block manifest and the template inventory a plan is built from, and `verifyTemplatePlanOutput` checks an output against the source, per block, without trusting the writer.
