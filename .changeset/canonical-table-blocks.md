---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

The experimental canonical session now runs the editor's insert table and delete table commands, in editing and suggesting mode. An inserted table comes with the empty paragraph the editor places after it; in suggesting mode its rows and paragraphs are recorded as inserted, and a deleted table's rows as deleted. The shared intent compiler gains a `tableBlock` intent whose inserted table takes fresh paragraph identities.
