---
"@stll/docx-core": minor
"@stll/folio-core": minor
---

Accepting a deleted paragraph mark that ends a section now joins the paragraphs across bookmark, comment or move range boundaries standing before the next paragraph, putting them where the resolved section ended, instead of refusing. A suggested copy of bookmark or move range boundaries keeps its copy ownership however many tracked insertions it is nested in. A suggested deletion over a bidirectional or other transparent wrapper holding a comment anchor whose range it does not remove whole is refused as holding a dependent record, as the bare anchor is.
