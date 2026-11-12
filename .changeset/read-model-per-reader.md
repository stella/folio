---
"@stll/folio-core": patch
"@stll/folio-react": patch
---

React's `PagedEditor` passes its document model on read-only: `HiddenProseMirror`, `NoteStoryEditor` and `HiddenHeaderFooterPMs` take a `ReadonlyDocument` as their `document` prop, and core's `createHiddenEditorState`, `runLayoutPipeline` and `waitForInitialLayoutFonts` take a model. Each reads it through `readModel` where it needs a plain `Document`.
