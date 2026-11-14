---
"@stll/folio-core": minor
"@stll/folio-react": minor
"@stll/folio-vue": minor
---

`DocxEditorRef.getPositionPage(position)` returns the rendered page that shows a body text position (story, paragraph id, offset) in both the canonical and the default session, and `FolioEditor.getPositionPage` does the same headlessly. `TextPosition`, `OpStory` and `OP_STORIES` are exported from `@stll/folio-core`. `PagedEditorRef.getPageNumberForPmPos` is removed; give `getPositionPage` a text position instead. React's `getTargetPage` reads the page from the layout, as Vue's does.
