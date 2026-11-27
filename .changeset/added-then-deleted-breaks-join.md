---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

A paragraph break one reviewer inserted and another deleted now reads as a pending join, as any deleted break does: a paragraph suggested before it copies the properties of the paragraph the join leaves, a merge that runs on through it gives the joined paragraph the first paragraph's properties, and a paragraph put before the join's last part goes before the whole join. Accepting such a batch gives every block the kind, style and list level the same batch gives it applied directly.
