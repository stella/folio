---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

A tracked batch that deletes the paragraph ending the body and inserts beside it now applies both operations in the experimental canonical session, as the default engine does. `DELETE_BLOCKS` with a revision now removes a paragraph mark that the same author's pending insertion added, joining that paragraph to the next one, instead of refusing the deletion.
