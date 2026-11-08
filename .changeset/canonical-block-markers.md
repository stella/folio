---
"@stll/folio-core": patch
---

The experimental canonical session now opens a document with a bookmark or a move range that starts or ends between paragraphs. Before, the whole document was refused. Typing and Enter beside such a marker keep it where it was through save. Joining two paragraphs across one still refuses and leaves the document unchanged.
