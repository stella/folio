---
"@stll/docx-core": patch
"@stll/folio-core": minor
"@stll/folio-agents": minor
---

Keep `listRendering.levelStarts` through a Document → ProseMirror → Document rebuild so custom list starts render without a DOCX round-trip, and add `formattingScope: "allParagraphs"` to block insertions so a multiline `text` can produce several list items. `ListRendering.levelStarts`, `DocumentSettings.mirrorMargins`, and the header/footer verbatim capture fields are now declared on the model types instead of attached through local intersections.
