---
"@stll/folio-core": minor
---

Keep a drawing's authored EMUs across the editor projection. Sizes, stroke
widths, wrap insets and text-box margins were measured into pixels for the
editor and converted back on save, and neither conversion is exact, so opening
a document and saving it again moved every image, shape and text box off the
numbers its author wrote. Each of the three nodes now carries the authored EMU
beside the pixels it was projected into (`_docxAuthoredEmu` on `ImageAttrs`,
`ShapeAttrs` and `TextBoxAttrs`) and writes it back while the pixel attribute
still projects from it; a command that moves the pixels still reaches the
document.

Three defaults went with it, because each was written back as a value the
document never had: the shape node's `outlineWidth` default of `1` gave an
`a:ln` with no `@w` a width of 9525 EMU, the text-box node's margin defaults
gave a text box with no `w:bodyPr` insets four authored ones, and a text box's
authored inset of zero was dropped by a truthiness test. Every consumer already
resolves an absent value against its own default.
