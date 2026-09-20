---
"@stll/docx-core": minor
"@stll/folio-core": patch
---

Type a shape or text-box outline's dash as `ST_PresetLineDashVal`, and give each stroke vocabulary its own table.

`ShapeOutline.style` was a hand-written eleven-member union named after CSS, holding what `a:ln/a:prstDash@val` declares. It is now `ShapeOutline.dash`, typed with `PresetLineDashVal`, generated from the committed schema graph by `scripts/generate-preset-line-dash.ts` with the same write/check pair `BorderStyle` uses; `bun run generate:preset-line-dash:check` runs in CI. A `@val` the schema does not declare is kept as `{ kind: "unrecognised", raw }`, written back unchanged, and reported through `ParseContext` as `outline-dash-outside-enum`. `a:custDash` is a different element and stays unmodelled: an outline that carries one replays through `ShapeOutline.rawXml`.

The display list resolved three vocabularies through one lookup keyed by lower-cased strings: a CSS `border-style`, a DrawingML preset dash, and a CSS `text-decoration-style`. `dash`, `dot` and `solid` collide across them, and every member no other vocabulary spells the same way had no entry and painted as a plain line. Nine of the eleven preset dashes (`dot`, `lgDash`, `dashDot`, `lgDashDot`, `lgDashDotDot`, `sysDash`, `sysDot`, `sysDashDot`, `sysDashDotDot`) and the seven heavy underline members were in that set, so a `sysDash` outline and a `dottedHeavy` underline both stroked solid. There are now three tables, each `as const satisfies Record<Union, StrokePattern>` over its own vocabulary, and each consumer calls the one it speaks.

The DOM painter had the same defect one step further on: it interpolated the outline's dash straight into a CSS `border` shorthand, so `border: 2px sysDash #000` was invalid and a dashed text-box outline did not paint at all. A dash is now translated to a CSS keyword before it reaches a shorthand, and `run.underline.style` is translated rather than assigned, which is what made `text-decoration-style: dottedHeavy` a no-op.
