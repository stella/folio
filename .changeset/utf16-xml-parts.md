---
"@stll/folio-core": patch
---

The conformance check reads an XML part written in UTF-16 (as word processors often write custom XML items) in that encoding instead of reporting it as not well formed.
