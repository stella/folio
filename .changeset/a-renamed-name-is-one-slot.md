---
"@stll/folio-core": patch
---

Read and write the names ECMA-376 Part 1 spells by writing direction from one generated table.

Part 1 names a horizontal edge `start`/`end` where Part 4 names it `left`/`right`, and Part 4 declares both so a Transitional consumer reads either. folio rebuilds every package as Transitional, so each of the readers, the writers and the survival law had its own hand-written copy of which two names are one slot.

`packages/core/src/docx/strictNames.gen.ts` is now that table, derived from the committed schema graph and a cited list of which spelling Part 1 declares. The border, cell-margin and indent readers take both spellings through it rather than falling back name by name, the table and numbering serializers bind the name they write to the key it stands for, and a check fails when a reader hand-lists a spelling again. No output changes: the same equivalences, now generated.
