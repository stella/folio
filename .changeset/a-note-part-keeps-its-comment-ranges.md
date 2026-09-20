---
"@stll/folio-core": patch
---

Refuse a note-part patch that would leave a comment range with only one half.
A comment can be anchored on a footnote's or endnote's own text, so its range
spans that note's paragraphs; splicing only the paragraph an edit touched then
wrote the other half alone, which is invalid OOXML and anchors the comment to
nothing. The refusal now belongs to the one splice primitive every selective
patch goes through, and a refused note part is rewritten whole from the model
instead of failing the save.
