---
"@stll/docx-core": patch
---

Track the deletion of a paragraph right before a table whose rows are still a pending insertion as its words and mark: accepting removes the emptied paragraph, rejecting restores it. A paragraph the same author added whole goes outright, and a tracked join that would run a paragraph into such a table is refused.
