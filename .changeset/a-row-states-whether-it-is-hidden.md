---
"@stll/folio-core": minor
---

`w:trPr/w:hidden` is tri-state: absent, an explicit on, an explicit off.

`tableRow`'s `hidden` attr defaulted to `false`, so the editor could not tell a row that authored `<w:hidden w:val="0"/>` from one that authored nothing, and the parser had to let an explicit off travel as captured bytes rather than model it. The default is `null` now, as `heightRule`'s already is, and the parser reads all three states into `TableRowFormatting.hidden`.

The attr-schema version moves to 5: a version-4 snapshot's `false` never meant an explicit off, so keeping it as one would start writing an element the document never carried, and the step drops it.
