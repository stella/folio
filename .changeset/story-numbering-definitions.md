---
"@stll/folio-core": minor
---

`FolioDocxReviewer.readNumberingDefinitions(story)` reads the numbering of a header, footer or note story, including the lists that story's operations started, which label its paragraphs. Called without a story it reads the main story as before. It returns `null` for a story the document does not have.
