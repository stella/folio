---
"@stll/folio-core": minor
"@stll/folio-react": minor
"@stll/folio-vue": minor
---

Add `highlightPassage` / `clearPassageHighlight` to `DocxEditorRef`: resolve `{ blockId, text }` to a range inside the block, scroll to it, and paint a persistent translucent passage highlight, falling back to scroll-to-block with a paragraph flash when the text no longer matches. Core exports the framework-neutral `resolvePassageRange` resolver.
