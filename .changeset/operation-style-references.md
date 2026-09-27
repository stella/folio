---
"@stll/folio-core": patch
"@stll/folio-agents": patch
---

Document operations now refuse a `styleId` that is not a paragraph style the document defines (an unknown id, or a table or character style): the operation is skipped with the new `missingStyle` reason before anything is applied, in every mode and through `suggest_changes`, instead of reporting a restyle that saves with no effect. `compareDocx` and `generateRedlineDocx` still carry the revised document's style references as it holds them.
