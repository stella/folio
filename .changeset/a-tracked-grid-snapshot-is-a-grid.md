---
"@stll/docx-core": minor
"@stll/folio-core": minor
---

Model `w:tblGridChange` as the grid it records rather than as the bytes it arrived in.

A `w:tblGridChange` holds a `w:tblGrid` of its own, and that grid holds its own `w:gridCol` children: a container nested in one of its own kind. It travelled as `TableFormatting.gridChangeXml`, a verbatim slot, so the snapshot's grid and every column in it existed only as markup nothing could read, and a rebuild could only copy the string back. The public corpus has the element in 28 packages, 104 columns in all, so this is a shape documents actually carry.

`TableFormatting.gridChangeXml` is replaced by `TableFormatting.gridChange`, a `TableGridChange` holding the revision's `@w:id` and one entry per `w:gridCol`. `w:w` is optional on a `w:gridCol`, so a column the snapshot stated no width for is `undefined` rather than zero, and it is written back without a width: a snapshot is a record of what stood, and a column with no measure is not a column of width zero.
