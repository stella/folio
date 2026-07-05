---
"@stll/folio-core": minor
"@stll/folio-vue": minor
"@stll/folio-react": patch
---

Add `@stll/folio-vue`, a Vue 3 editor adapter over `@stll/folio-core` that tracks
the `@stll/folio-react` editor contract (`DocxEditor`, `DocxEditorProps`,
`DocxEditorRef`, `renderAsync`).

To share one framework-neutral base across adapters, the folio UI translation
catalog moves into `@stll/folio-core/i18n/messages` (the React `messages` subpath
re-exports it, so `@stll/folio-react/messages` is unchanged). Core also gains the
helpers the adapters build on: `ClipboardManager`, `AutoSaveManager`,
`resolveColorToHex`, the `WrapType` union, `docx` + `prosemirror/extensions`
barrels, and a set of ported editor-engine helpers (comment/section-break/table/
image commands, tracked-change extraction, visual-line navigation, image layout).
