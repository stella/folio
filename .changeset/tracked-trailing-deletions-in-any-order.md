---
"@stll/docx-core": patch
---

Tracked deletions of a container's last paragraphs accept as the same deletions made directly in any order. Deleting a paragraph whose break a pending deletion already takes, into paragraphs that keep nothing once accepted up to the container's end, now also records the break before it deleted, as when that paragraph is last; accepting no longer leaves an empty paragraph behind.
