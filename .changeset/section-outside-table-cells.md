---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

Document operations keep section breaks out of table cells. An operation after which a section ends in a cell where none ended before is refused with `sectionBoundary`, whether it makes, moves, copies or restores the break, and so is setting a section endpoint on a cell's paragraph. A break a document already held in a cell stays through edits that leave it where it is. A paste or a drop into a cell, and a pasted table, come without the section breaks of their paragraphs.
