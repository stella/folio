---
"@stll/docx-core": patch
---

Delete a paragraph that ends a section, tracked or directly. The section break is the paragraph mark's: a direct deletion, or accepting the tracked one, joins the sections on either side under the properties of the section after it; rejecting restores every section. One's own added mark that ends a section records the deletion beside the insertion instead of taking the insertion back.
