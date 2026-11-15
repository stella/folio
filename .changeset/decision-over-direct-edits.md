---
"@stll/docx-core": patch
---

A revision decision (accepting or rejecting a tracked change) passes a concurrent direct positional edit, such as typing, a deletion or a format, unchanged, rather than being refused. It is refused only when the edit cut a revision record in two and gave its second half a new id. A reviewer's decision is no longer dropped because another pending edit of the same peer, unrelated to it, was refused.
