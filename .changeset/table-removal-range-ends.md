---
"@stll/docx-core": patch
---

Removing a table row or a whole table keeps the end of a bookmark or move range that reaches past it where the row or table was, so the document still saves with every range whole.
