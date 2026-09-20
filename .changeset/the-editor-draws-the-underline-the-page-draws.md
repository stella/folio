---
"@stll/folio-core": patch
---

Render an underline in the editor from the same table the painters read.

`UnderlineExtension.toDOM` carried a private four-entry table (`double`,
`dotted`, `dash`, `wave`) while the display list and the DOM painter read the
total one, so the eleven members it had never heard of drew a plain line in the
editor and their own pattern on the page: `dottedHeavy`, `dashLongHeavy`,
`wavyHeavy` and `words` among them. The table is gone. `toDOM`, the DOM painter
and `textToStyle` now all render a member through `underlineDecorationCss`,
which is `as const satisfies Record<UnderlineStyle, …>`, so a member added to
the enumeration cannot reach a backend without a decision attached.

The one table also states the weight CSS has no keyword for: `thick` and the
seven `*Heavy` members carry a `text-decoration-thickness` of twice the ratio
the display list strokes a plain underline with, so both DOM backends scale it
with the font size. The remaining approximations are recorded beside the table:
`words` underlines the spaces between words (`text-decoration-skip-ink` skips
descender ink, not spaces, and `text-decoration-skip: spaces` never shipped),
`wavyDouble` draws two straight lines, and `dashLong`, `dotDash` and
`dotDotDash` draw the single dash pattern CSS has.

`toDOM` states the line and the style in one `text-decoration` shorthand, which
the mark's existing parse rule reads back through the table's inverse: each CSS
keyword resolves to its canonical author, the plain member drawn exactly that
way (`solid` to `single`, `dashed` to `dash`, `wavy` to `wave`). The table is
many-to-one, so what CSS cannot spell does not survive the round trip:
`dottedHeavy` parses back as `dotted`, `words` and `thick` as `single`. A value
naming no keyword folio writes parses as the plain underline, and a
`text-decoration-style` is not read on its own, which would make a dotted
strikethrough an underline.

`w:u w:val="none"` now cancels an inherited underline in the DOM painter as it
already did in the display list: the member carries no declarations, so the
painter draws no line rather than an unstyled one.
