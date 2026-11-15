---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

Paragraph properties an edit sets read in the live model as the saved file reopens: the spacing sides the paragraph states are marked as the parser marks them, and a hanging indent is held as the parser reads it, a negative first line flagged hanging. `paragraphSpacingExplicit` is the one rule for the spacing sides, used by the parser and by the patch an operation states; paragraphs no operation names keep their formatting as it is.

A tracked paragraph property change records the properties the paragraph now has as its current formatting, as a reopened file reads them from the paragraph itself; folding a later change keeps the first change's previous properties.
