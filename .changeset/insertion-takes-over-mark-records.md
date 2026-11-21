---
"@stll/docx-core": patch
---

A tracked paragraph put after one that stands before a table whose mark it takes over now takes that paragraph's pending property change and inserted break with the mark, as a tracked split's second half does, instead of being refused. Accepting reads as the same insertion made directly; rejecting the insertion gives the paragraph back with its records. A break pending deletion or moving away, and a new final paragraph of its container, stay refused.
