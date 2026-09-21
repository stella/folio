---
"@stll/folio-core": patch
---

Align a cell's paragraphs on each side of its nested table separately.

A cell's paragraphs were aligned as one sequence, so a cell that lost paragraphs could pair a surviving one with a paragraph on the far side of its nested table. Nothing moves a block across a table, and the cell's last paragraph cannot be deleted because its mark has nothing to join, so the redline kept a paragraph the revised document does not have and the comparison refused its own round trip with `container`. The paragraphs between a cell's nested tables are now aligned run by run, the way rows and cells already are.

A cell whose last block is a table states a cell no consumer renders as written: `CT_Tc` ends in a paragraph, and every consumer reads the implied empty one there. Producers leave that shape behind when they rewrite a cell and delete the closing paragraph along with the rest. Parsing it as written asked the comparison to delete a paragraph mark that has to stay, so the parser now reads the implied paragraph as the fact it is, the way it already supplies one for a cell that states no block at all.
