---
"@stll/docx-core": minor
"@stll/folio-core": minor
---

Pair `ST_TextDirection`'s two spellings of each flow by ECMA-376 Part 4
§14.11.7. Folio paired them by the letters in the token, so every one of the
six Strict spellings rendered as something other than its Transitional twin:
`tb` is the horizontal flow and turned a quarter clockwise, `rl` and `lr` are
vertical flows and painted flat. Rendering is now decided per flow, and the
section's own text direction is narrowed against the enumeration rather than a
second hand-written copy of it.
