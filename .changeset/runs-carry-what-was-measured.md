---
"@stll/folio-core": minor
---

Carry what the measurer applied on the glyph run, so a backend can reproduce
the width the line was fitted at.

`advancesPx` folded letter spacing, a horizontal scale, justification, kerning
and small capitals into one number per code point. A backend that hands the run
to a shaper cannot recover any of them from the text: the shaper advances glyphs
by what the font says, and none of those five is in the font. The DOM backend
therefore painted runs at the glyphs' own width rather than the laid-out one,
by as much as 78 px on a justified line.

`DisplayGlyphRun` now names them: `adjustments` carries the letter spacing,
horizontal scale and per-space justification delta, and `kerning` and
`smallCaps` state what the advances were measured with. `DisplayFontFace` gains
`fallbacks`, the families between the first and the generic, because a face is a
stack and a backend handed only its first entry paints a different face from the
one measured wherever that entry is missing.

A run's advances now also sum to the width the line was broken on. They were the
sum of per-character measurements, which differs from the string's own width
wherever a pair kerns or ligates; the difference is spread across the run rather
than left as an extent no backend paints.

The editor's run-drift check gates the painted extent as well as the origin:
every run in the two fixtures now lands within one browser layout quantum.
