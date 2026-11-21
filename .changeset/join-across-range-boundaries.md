---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

Join two paragraphs that only range boundaries standing between blocks separate: a direct join takes the boundaries in where the paragraphs meet and its undo stands them between the paragraphs again; a tracked join leaves them in place until it resolves. The canonical session merges such paragraphs in direct and tracked batches.
