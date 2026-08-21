---
"@stll/folio-core": minor
"@stll/docx-core": minor
---

Headless DOCX report generation: `/server` exports typed builders (`heading`, `paragraph`, `run`, `table`, `pageBreak`, `hyperlink`, `bookmark`, `endnote`, `createTableOfContentsField`); external hyperlinks inside headers, footers, footnotes and endnotes now get relationships in their own rels part; `createEmptyDocument` initialises `package.relationships` so in-memory headers and footers materialise; `DocumentSettings.updateFields` round-trips; the Stella style set gains `Heading1`-`Heading6`, `TOCHeading`, `TOC1`-`TOC3`, `EndnoteReference` and `EndnoteText`; complex fields keep `w:dirty`/`w:fldLock` across parse and save.
