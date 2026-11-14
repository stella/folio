---
"@stll/docx-core": patch
---

A paragraph replacement, inline replacement, paragraph review or container replacement now writes every revision its new content holds, so an undo that puts a tracked change back is refused with `partConflict` against a concurrent decision of that change instead of the two orders diverging.
