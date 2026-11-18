---
"@stll/docx-core": patch
---

Refuse a tracked range deletion that would hold only empty hyperlinks, content controls or wrappers, and plan none for them in a suggestion: no run carries such a record, so no file or reader showed it. A suggested removal of a paragraph holding only an empty link now records just the join, or is refused when there is nothing to record.
