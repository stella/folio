---
"@stll/docx-core": patch
---

A tracked split whose new paragraph is the second half now refuses mark run properties other than the paragraph's own. That half ends with the paragraph's own mark, and no property change records its run properties, so rejecting the split could not give them back.
