---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

A tracked or suggested deletion of a paragraph one's own pending insertion added takes the insertion back: before a following paragraph it is removed outright, and after the container's last paragraph the insertion is rejected, so the paragraph before keeps its identity and properties. The legacy session no longer leaves an empty paragraph-property change on that paragraph.
