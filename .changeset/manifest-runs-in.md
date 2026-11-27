---
"@stll/folio-core": minor
---

The block manifest reports `runsIn` on a paragraph whose mark is always hidden (`w:specVanish` in the mark's run properties, a style separator): the paragraph runs on into the one after it, on the same line, as a heading run into its text. A caller can join the two into one paragraph of the output, which keeps no hidden mark.
