---
"@stll/folio-core": patch
---

A `w:numPr` naming numbering the file never defines no longer fails a save. Parsing and style-set extraction rewrite such a reference to the `w:numId w:val="0"` sentinel on both the paragraph and the paragraph style, with a warning naming the style, so a style whose numbering is missing stays unnumbered instead of inheriting its `w:basedOn` parent's list; deleting the `w:numPr`, which the paragraph tier did before, numbered a paragraph its source showed unnumbered. Building a package now checks the numbering a style names, not every `w:num` the source carried.
