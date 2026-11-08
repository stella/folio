---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

A commented public operation on a paragraph inside a table cell or a block content control now applies in the experimental canonical session. It no longer refuses with "The comment anchor no longer exists." A commented insertion beside such a paragraph now carries its comment too. `storyParagraphs` is exported from `@stll/docx-core/ops`.
