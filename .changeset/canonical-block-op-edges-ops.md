---
"@stll/docx-core": patch
---

Tracked paragraph insertions and deletions next to tables, block-level range boundaries and a container's final paragraph now resolve one record at a time without reminting the anchor's record ids, and a tracked deletion before a table that cannot take its mark is refused up front. Direct joins and paragraph deletions now leave the caret in a paragraph the edit keeps.
