---
"@stll/docx-core": patch
---

Every operation kind except comment operations can be sent in a batch: tables, numbering, sections, package resources, headers and footers, notes, paragraph and inline replacement, paragraph review and their undo inverses, with bounded validation of the model records they carry. Operations without a position transform are sequenced by footprint: they pass concurrent edits of other paragraphs, tables and package parts, and an overlap is refused with a typed reason (`blockConflict`, `tableRequiresExclusiveEdit` or `partConflict` naming one of `EXCLUSIVE_PARTS`). Text typed at a paragraph's start and a concurrent format from the start now commute.
