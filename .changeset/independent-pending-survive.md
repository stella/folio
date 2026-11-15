---
"@stll/docx-core": patch
---

The optimistic client no longer drops every pending batch after one it drops. When a pending batch does not rebase over a sequenced one, or is rejected, the batches after it rebase as if it had never been made, over its inverse first, and only a batch that depends on it is dropped with it.
