---
"@stll/folio-core": patch
---

Apply tracked-changes and suggested `replaceInBlock`, `replaceRange` and `replaceBlock` edits as their redline in paragraphs with tabs, breaks, fields, content controls, note references or drawings too: each change marks only the characters it removes as a deletion and inserts only its new text, instead of deleting and reinserting the whole match. Inline content outside the changed characters carries no revision. Background clearing records a run-property change only on the kept characters of a highlighted stretch a change touches; removed characters keep their highlight as deleted text, and the new text is written without it.
