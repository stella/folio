---
"@stll/docx-core": patch
"@stll/folio-core": patch
"@stll/folio-react": patch
"@stll/folio-vue": patch
---

The style resolver and the style engine take style definitions typed read-only and hand out read-only results: `getStyle`, `getParagraphStyles`, `getTableStyles`, `getDocDefaults`, the default style getters, `resolveParagraphStyle`, `resolveParagraphStyleInTable`, `resolveRunStyle` and `getRunStyleOwnProperties`. `getRunStyleOwnProperties` returns the style's own run properties instead of a shallow copy. The text and paragraph formatting merges keep plain inputs plain and give read-only results for read-only inputs, and a merge with a source is a new record. The paragraph attributes the style cascade fills (`borders`, `shading`, `tabs`, `defaultTextFormatting`, `_tableRunFormatting`, `_resolvedFormatting`, `_styleResolvedFormatting`), a table cell's `borders` and `_resolvedBorders`, and the selection snapshot's `textFormatting` and `paragraphFormatting` are read-only. `writableCopy` in `@stll/folio-core/types/documentModel` is the copy a writer takes before it changes such a value. `DeepReadonly` moves to `@stll/docx-core/model`, and core re-exports it. React's `PagedEditor` reads the model's styles without `readModel` and passes them on read-only; `getParagraphTabs` and React's and Vue's ruler tab stops are read-only. Saving a paragraph's or a table cell's borders, shading and tab stops copies them, so the saved document owns what it states.
