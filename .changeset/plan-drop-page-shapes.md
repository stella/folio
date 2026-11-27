---
"@stll/folio-core": minor
---

A template plan can write a paragraph without its drawn shapes that are fixed to the page (`dropPageShapes` on a source item and on a nested paragraph): single floating shapes without a picture and without text of their own whose horizontal place is counted from the page or a margin, as a conversion from a fixed layout leaves them behind. The block manifest counts them (`pageShapes`), verification holds the output to what the plan names, and the report counts what was left out.
