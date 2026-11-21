---
"@stll/docx-core": minor
"@stll/folio-core": minor
"@stll/folio-agents": patch
---

Refuse a table operation over a table, row, cell or grid that already carries a pending change with its own reason: docx-core's `pendingTableChange` (was `revisionConflict`), reported as the new `pendingTableChange` skip reason. The legacy engine reports it too where it refuses the same shape: a tracked deletion of a row, table or column whose rows or cells are already pending insertion or deletion.
