---
"@stll/folio-vue": minor
---

Wire `DocxEditorRef.getEditorRef()` on the Vue adapter, moving it from
deferred to paired with React's `PagedEditorRef`. The Vue package has no
ported `PagedEditor` sub-component to source a handle from, so
`useDocxEditorRefApi.ts` synthesizes an equivalent `PagedEditorRef`-shaped
object from primitives it already holds: `getEditor` / `getDocument` /
`getState` / `getView` / `ensureView` / `focus` / `blur` / `isFocused` /
`dispatch` / `undo` / `redo` / `canUndo` / `canRedo` / `setSelection` /
`getLayout` / `relayout` delegate to the headless `FolioEditor` controller;
`scrollToPosition` reuses the pages-scroll helper; `scrollToPage` /
`scrollToParaId` reuse the existing top-level ref implementations;
`getPageNumberForPmPos` resolves through the same layout helper
`getCurrentPage` already uses. `getHfView` stays a documented no-op (returns
`null`) since Vue has no persistent hidden header/footer `EditorView` yet.
