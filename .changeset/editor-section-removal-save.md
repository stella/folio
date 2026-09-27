---
"@stll/folio-core": patch
"@stll/folio-react": patch
"@stll/folio-vue": patch
---

The React and Vue editors' save no longer throws after a section-ending paragraph is removed on purpose (a direct `deleteBlock` through the editor ref, or accepting its tracked deletion): the full repack is told which sections the editor's edits removed, as the headless reviewer's save already was.
