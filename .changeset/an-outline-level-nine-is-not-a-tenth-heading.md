---
"@stll/docx-core": minor
"@stll/folio-core": minor
---

Model `w:outlineLvl` as a union, so the reserved body-text value cannot be read as a tenth heading level.

`ParagraphFormatting.outlineLevel` was `number`, which put ECMA-376 17.3.1.20's rule ("9 specifically indicates that there is no outline level applied to this paragraph") in every consumer's hands. `OutlineLevel` is now `{ kind: "bodyText" } | { kind: "heading"; level: 0..8 }`, with `level` a union of nine literal types: the sentinel has no representation as a heading, an out-of-range value has none at all, and an absent field still means "states none, inherits one".

One reader owns the parse boundary (`outlineLevelFromStatedValue`) and one writer owns the emit (`outlineLevelStatedValue`). The paragraph parser, the style parser, the style cascade, the display-list outline, the layout bridge, the ProseMirror attr and its validator, markdown, the style sets and the legal-source compiler all move to the union; `isHeadingOutlineLevel` and the bare `BODY_TEXT_OUTLINE_LEVEL = 9` are gone, replaced by `headingLevelOf` and the body-text arm.

A `w:outlineLvl` outside 0..9 is now dropped at the parse boundary rather than carried through the model, which is what the Rust projection kernel already did. The container-survival census records the one value that stops surviving a rebuild.
