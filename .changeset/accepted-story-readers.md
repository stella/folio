---
"@stll/folio-core": minor
"@stll/folio-react": patch
"@stll/folio-vue": patch
---

Resolve pending block deletions and paragraph joins consistently in clean Markdown and plain-text readers.

Import these moved exports from `@stll/folio-core/docx/storyPlainText`:

- `getDocumentText`, `getWordCount`, and `getCharacterCount` (previously `@stll/folio-core/docx/documentParser`).
- `getTableText` (previously `@stll/folio-core/docx/tableParser`).
- `getTextBoxText` (previously `@stll/folio-core/docx/textBoxParser`).
- `getHeaderFooterText` (previously `@stll/folio-core/docx/headerFooterParser`).
- `getFootnoteText` and `getEndnoteText` (previously `@stll/folio-core/docx/footnoteParser`).
