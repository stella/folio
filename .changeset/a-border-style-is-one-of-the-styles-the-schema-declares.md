---
"@stll/docx-core": minor
"@stll/folio-core": minor
"@stll/folio-react": minor
"@stll/folio-vue": minor
---

Type `BorderSpec.style` as the `ST_Border` enumeration instead of `string`.

`w:val` on `CT_Border` has 193 members. The model held it as a bare `string`, beside a hand-written 22-member `KnownBorderStyle` the parser narrowed against; the hand list omitted every page-border art glyph and four line styles, so a document that used one reached the consumers as a value no rendering table knew, and nothing could say which table was missing which member. `BorderStyle` is now generated from the committed schema graph by `bun run generate:border-styles`, and a schema refresh that adds a member fails the generator check rather than widening a `string`.

`nil` and `none` stay distinct members: `none` cancels a border inherited from the container, `nil` states that none is set, and Word round-trips whichever the author wrote. Consumers ask `statesNoBorder`, `isBorderNone` or `isBorderNil` rather than comparing the token; `specifications/reserved-values` records the decision and the lint holds it.

A `w:val` the schema does not declare is kept verbatim as `{ kind: "unrecognised", raw }` and written back unchanged, with a `border-style-outside-enum` parse warning so the normalisation is visible. Refusing it would drop an edge Word paints, and reading it as a default would rewrite the document on open.

One table now says how a member renders. There were three — the layout bridge's, `formatToStyle`'s and `TableExtension`'s — and they covered different amounts of the enumeration, so a `thinThickSmallGap` cell edge came out `double` in the editor and `solid` on the paginated page, and a `dotDash` paragraph rule came out `dashed` through the bridge and `solid` through `borderToStyle`. `CSS_BORDER_STYLES` is total over the union at compile time, and a page border now takes the 3px floor for every member that paints as a CSS `double`, not only for `w:val="double"`.

`KnownBorderStyle` is removed; `BorderStyle`, `BorderStyleValue` and `UnrecognisedBorderStyle` replace it. The layout engine's `BorderStyle.style` and `CellBorderSpec.style` are typed `CssBorderStyle`, which is what they always held, so a measurer or painter can no longer ask whether a laid-out border is `"nil"`. `TableCellBorderCommandSpec` and `TableBorderCommandSpec` are exported from `@stll/folio-core/prosemirror` and carry the same union, and both adapters' table-style presets use them instead of re-declaring the shape.
