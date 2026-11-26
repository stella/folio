---
"@stll/folio-core": minor
"@stll/docx-core": patch
"@stll/folio-react": patch
"@stll/folio-vue": patch
---

Use canonical review for tracked fidelity, comment receipts, and saved suggestion positions. Pass a canonical document model to `createPendingSuggestionSourceSnapshot` and a saved source callback to `placeRemovedAnchors`. The deep ProseMirror fidelity and snapshot-only comment-mark helpers are test-only or removed.
