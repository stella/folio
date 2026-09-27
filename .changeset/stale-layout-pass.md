---
"@stll/folio-core": minor
"@stll/folio-react": patch
"@stll/folio-vue": patch
---

A scheduled layout pass lays out the editor's state when it runs, so a pass an edit scheduled before `loadDocument` no longer paints the replaced document's text into the loaded one; an incremental pass derives what to re-measure from the committed and current documents. `createLayoutScheduler` takes `readState` and `schedule()` no longer takes a state or dirty range.
