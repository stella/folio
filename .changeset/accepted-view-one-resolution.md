---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

`resolvedDocumentView` resolves every revision of every story one way, as Accept All or Reject All does. Clean Markdown and the block plain-text reader read their accepted view from it, so they keep the same paragraphs, tables and notes a reviewer keeps after accepting all.
