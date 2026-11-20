---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

A paragraph mark can hold another author's pending insertion and a later pending deletion of it. Both records are read, written in schema order, listed as changes and resolved one at a time or together; deleting another author's inserted paragraph records the deletion beside the insertion instead of removing the paragraph.
