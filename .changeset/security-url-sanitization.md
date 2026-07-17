---
"@stll/folio-core": patch
"@stll/folio-react": patch
"@stll/folio-vue": patch
---

Sanitize hyperlink and image link URLs so pasted, programmatic, or DOCX-sourced
`javascript:`/`data:`/`file:` targets can no longer reach the live DOM or be
opened. Hyperlink marks are sanitized on parse (`parseDOM`), on render
(`toDOM`), and when set/inserted/edited; image `a:hlinkClick` targets are
sanitized at parse time; the Vue popup `window.open` path now mirrors React's
sanitizer; and aux-click on link anchors no longer bypasses the guard. Internal
bookmark anchors (`#name`) are preserved.
