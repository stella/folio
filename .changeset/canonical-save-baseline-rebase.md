---
"@stll/folio-core": patch
"@stll/folio-react": patch
"@stll/folio-vue": patch
---

After a canonical save, the editor makes the saved package the baseline later saves patch against, so a text edit after Enter is written as a paragraph patch of the last saved file again instead of a structural splice or a full rewrite. An undo past the save counts as a change like any other edit. `markCanonicalSaved` takes the saved package for this.
