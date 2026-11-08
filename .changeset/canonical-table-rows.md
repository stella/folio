---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

The experimental canonical session now runs the editor's row commands: insert a row above or below and delete rows, in editing and suggesting mode. Each saves what the editor would save and undoes exactly; deleting every row removes the table. The shared intent compiler gains a `tableRow` intent whose inserted row takes fresh paragraph identities.
