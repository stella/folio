---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

Keep `w:tblGridChange` and `w:numberingChange` when the container they sit in is rebuilt. Both are tracked-change records — the grid a reviewer replaced when resizing a column, and the numbering a reviewer replaced when changing a list — and nothing in the editable model derives either, so a save that rebuilt the container dropped the revision and the document then read as though the change had always been there. They now travel as their own capture slots (`TableFormatting.gridChangeXml`, `ParagraphFormatting.numberingChangeXml`) and are written back on both the replay and the rebuild path. A captured `w:pPr` carrying a `w:numberingChange` is no longer refused for replay either; refusing it used to force the rebuild that could not write it.
