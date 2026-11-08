---
"@stll/folio-core": patch
---

A save that rewrites a paragraph with a style keeps the run properties its runs share with the paragraph mark (such as `w:position` or `w:sz`), and no longer writes properties turned off on its runs for the mark's sake. This covers edits, accepted and rejected changes, and typing in the editor.
