---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

Deleting a paragraph beside a numbered one now leaves the same list numbering whether it is applied directly or tracked and then accepted: a container's last paragraph under a pending deletion of its mark takes the properties the direct deletion gives it, a direct deletion joins the paragraph across block-level bookmarks as the tracked one does, and a tracked deletion before a table of a paragraph whose mark is another reviewer's pending insertion is recorded beside that insertion, so accepting both removes the paragraph whole.
