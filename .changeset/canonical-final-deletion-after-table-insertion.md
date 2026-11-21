---
"@stll/folio-core": patch
---

A canonical batch that deletes a container's last paragraph after a table, and puts a paragraph after that table, now applies both: the deletion runs after the insertion and removes the paragraph as it does one operation at a time, instead of being skipped.
