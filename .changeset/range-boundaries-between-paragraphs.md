---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

Read range boundaries standing between blocks as positions, not blocks: deleting a paragraph finds the paragraph or table past them, or ends its container, and a paragraph that joins the next one takes them in where the two meet. Resolving a paragraph mark before a table or block content control past them removes or keeps the paragraph as it would without them.
