---
"@stll/folio-core": minor
---

Shape text in the scripts that need it, from one implementation both the
measurer and the PDF backend use.

Arabic, Hebrew, the Indic scripts and their neighbours do not select one glyph
per code point: a letter takes its form from the letters beside it, lam followed
by alef ligates, a Devanagari cluster reorders into a conjunct, a Hebrew point
hangs off the letter it belongs to. A new `stella-text-shaper` crate answers
that question over rustybuzz, and `packages/core/src/shaping` is the only way to
it, so a measurement in CSS pixels and a PDF text matrix scale the same glyph
ids and the same advances.

The headless measure provider now measures such a run by its clusters rather
than a code point at a time, and the PDF backend paints the glyphs shaping
chose, subsets them, and maps each back to the characters that formed it so a
ligature or a conjunct still extracts as text. `writePdf` is asynchronous as a
result, and no longer reports `unshaped` runs or refuses them under
`strictShapedScripts`: the runs it used to name are painted correctly.

The shaper is a separate WebAssembly artifact with its own size budget, fetched
the first time a document actually contains a run that needs it. A document in
Latin, Cyrillic or Greek never loads it.
