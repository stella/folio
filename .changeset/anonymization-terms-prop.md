---
"@stll/folio-core": minor
"@stll/folio-react": minor
"@stll/folio-vue": minor
---

Anonymization highlights work in the canonical session as in the default one. The `anonymizationTerms` prop sets the terms to highlight; the editor finds them in the shown text of every story (body, table cells, headers, footers, footnotes, endnotes) and keeps them in step with every edit, a peer's included, without touching the document, its undo history or its pending changes. After an edit it rescans only the paragraphs the edit replaced. `onAnonymizationMatchesChange` and the new `DocxEditorRef.getAnonymizationMatches()` report each occurrence as `{ start, end, label, canonical }` with text positions, and the callback fires only when the occurrences change. Text in a tracked deletion is matched, as before.

`setAnonymizationTermsMeta`, `anonymizationDecorationsKey`, `getAnonymizationMatches(state)` and the `AnonymizationMatch` / `AnonymizationTerm` types are removed: pass terms through the `anonymizationTerms` prop and read occurrences through the ref. The term type is `FolioAnonymizationTerm`, the occurrence type `FolioAnonymizationMatch`.
