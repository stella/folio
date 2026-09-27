---
"@stll/folio-core": patch
---

A batch that deletes a container's last paragraph and also inserts a block after it now refuses, as `overlappingOperation`, a `setBlockParagraphProperties` on that paragraph (or the insertion, when it comes last). The insertion means the deletion takes the paragraph's mark as well, so the properties used to be written to a paragraph that accepting the redline removes.
