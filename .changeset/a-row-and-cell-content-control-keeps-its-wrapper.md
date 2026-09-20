---
"@stll/docx-core": minor
"@stll/folio-core": minor
---

Keep a row-level and a cell-level content control. A `w:sdt` between a table
and its rows, or between a row and its cells, was unwrapped: the rows and cells
were spliced in and the control — its tag, alias, lock, data binding and
`w:sdtEndPr` — went on the floor, so a template whose repeating section or
bound cell folio merely opened and saved came back unbound.

`TableRow.contentControls` and `TableCell.contentControls` record the control
on each child it wrapped, outermost first, and the save re-opens one wrapper
per run of consecutive children that name the same control. The record is on
the children rather than between them because a table's children are rows and a
row's are cells, and neither has a node to spare for a wrapper that is not one;
it rides the ProseMirror row and cell nodes as an attr, so splitting or moving
one keeps it inside its control. A control over several rows, and a control
inside a control, both come back as they were written.

`SdtProperties.endProperties` models `w:sdtEndPr`, which folio held only as
captured bytes: every control it rebuilt — one an edit touched, one a full
repack wrote — lost its end mark and the run properties on it. This covers the
block and inline levels too.
