---
"@stll/folio-vue": minor
---

Wire the Vue title-bar MenuBar to the real `MenuDropdown` / `TableGridInline`
primitives (replacing the inline render-null stubs) and route the document-insert
flow end to end. `MenuBar` items now drive image / table / page-break /
table-of-contents insertion through `DocxEditor.vue`, which prefers the host
`onInsertImage` / `onInsertTable` / `onInsertPageBreak` / `onInsertTOC` props when
provided and otherwise falls back to the core view-level helpers
(`insertImageFromFile`, `insertTableInView`, `insertPageBreakInView`,
`insertTableOfContentsInView`). `showTableInsert` now gates the Insert > Table
menu item, mirroring the React toolbar.
