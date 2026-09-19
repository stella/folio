---
"@stll/folio-core": patch
---

Every style table folio authors names its built-ins with the `w:name` Word itself writes, from one constant map in `docx/builtInStyles.ts`. The generic style set wrote `Heading 1`…`Heading 4` where Word writes `heading 1`, and the Stella set wrote `Footnote Text`, `Footnote Reference`, `Endnote Text`, `Endnote Reference` and `Footer` where Word writes those lowercase. Word's own casing is not uniform, so the map records each spelling separately. The TOC-entry style matcher now uses the same name normaliser as the classifier, accepting `toc 1`, `TOC 1` and `TOC  1` alike.
