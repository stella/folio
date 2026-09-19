---
"@stll/folio-core": patch
---

A heading pasted from another application is pointed at the open document's own heading style instead of the English built-in id the schema's paste rule can produce. In a document whose heading styles are localized, or one whose style set stops at level four, that id resolved to nothing and the pasted heading lost the document's heading formatting. A paragraph is retargeted only when the document defines no style under the id it carries and does define one at that outline level.
