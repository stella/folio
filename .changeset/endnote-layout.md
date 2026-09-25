---
"@stll/docx-core": minor
"@stll/folio-core": minor
---

Lay out and paint endnotes: they paginate with the body after its last block (`w:pos="docEnd"`) or at the end of each section (`sectEnd`, honouring `w:noEndnote`), open with a separator rule and continue under a continuation separator, show the same number as their body reference, and open the note editor on double-click. Read `w:settings/w:endnotePr`.
