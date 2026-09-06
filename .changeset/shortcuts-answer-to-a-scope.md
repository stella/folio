---
"@stll/folio-core": minor
"@stll/folio-react": minor
"@stll/folio-vue": minor
---

Let a host say which key presses the editor's page-level shortcuts answer.

`DocxEditor` takes a `keyboardShortcuts` prop: `"document"` (the unchanged
default) answers every press on the page, `"editor"` answers only a press
landing inside the editor, and `"none"` binds no page-level listener at all.
A host that docks the editor beside its own panes keeps its own bindings and
opens the dialog through the new `DocxEditorRef.openFind` / `openReplace`,
which seed the search box from the current selection exactly as Cmd/Ctrl+F
does. The scope predicate is `isKeydownInShortcutScope` in
`@stll/folio-core/managers/editorShortcuts`; `useWheelZoom` takes the same
scope in place of its `enableKeyboardShortcuts` flag.
