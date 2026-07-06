---
"@stll/folio-vue": patch
---

Fix the Document Outline panel always showing "No headings found": `outlineHeadings`
was declared but never populated. Wire the existing `useOutlineSidebar` composable
into `DocxEditor.vue` so opening the outline (via the toggle button, toolbar, or File
menu) collects headings from the document, re-collects them on further edits while the
panel stays open, and clicking a heading scrolls the visible pages to it.

Fix Find/Replace having no way to open: `useKeyboardShortcuts` existed but was never
invoked from `DocxEditor.vue`, so Ctrl+F / Ctrl+H / Ctrl+K never opened their dialogs.
Invoke the composable from the editor shell, and add a Find & Replace entry to the
File menu as a discoverable, mouse-driven entry point alongside the keyboard shortcut.
