---
"@stll/folio-core": patch
---

The AI-facing snapshot is linear in block count again. It resolved each block's
position twice, and `ProseMirror`'s `resolve` scans a fragment from index 0, so
a flat document cost O(blocks^2); the walk now carries the path it is already
on. A 4,000-paragraph snapshot drops from 124ms to 24ms, and every reviewer
read built on it drops with it. Fixing it surfaced a second defect: text in a
table nested inside a hidden `w:trPr/w:hidden` row reached the snapshot,
because the check consulted the nearest row rather than every enclosing row.
