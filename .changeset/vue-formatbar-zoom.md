---
"@stll/folio-vue": minor
---

Add standalone `FormattingBar` and `ZoomControl` Vue components, matching the
React adapter's public surface. Both are controlled and own no editor view, so a
host can render them outside `<DocxEditor>`. `FormattingBar` (`FormattingBarProps`)
composes the existing Vue pickers into the minimal rail (undo/redo | style |
B/I/U | color | align/lists/indent) and emits `format` / `undo` / `redo`;
`ZoomControl` (`ZoomControlProps`, `ZoomLevel`) is a compact zoom-level dropdown.
This closes the corresponding React/Vue export-parity divergence entries.
