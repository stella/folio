---
"@stll/folio-core": patch
---

Deleting a paragraph that ends a section (`deleteBlock` in direct mode, or a direct `mergeBlockWithNext` into or out of it) no longer leaves a reviewer whose save throws. The section break goes with the paragraph's mark and the content before it joins the following section, as accepting the same tracked edit already did; a direct merge now keeps the second paragraph's section break rather than the first's. A section removal accepted or applied earlier also stays saveable after later edits move the remaining section breaks.
