---
"@stll/docx-core": minor
"@stll/folio-core": minor
---

Carry a transparent inline wrapper into the editor instead of dropping what it said.

`w:bdo` and `w:dir` reached the editor as their content and nothing else: the projection flattened them, so opening a document lost the direction the author wrote and the painter drew the text in the paragraph's direction. The tree is still flattened — the inline loops narrow by a chain of `else if`, and a wrapper left in one would reach whichever branch happens to be last — but what the wrapper said now rides an `inlineWrapper` mark on the leaves it held. The mark's `stack` attr lists the wrappers a leaf sits inside, outermost first, because ProseMirror's mark set is unordered across types and two marks could not say which wrapper is inside which. `RunFormatting` gains `bidiWrapper`, the painter writes `unicode-bidi` and `dir` from it, and a glyph run takes the wrapper's direction over the paragraph's.

`BidiWrapper` becomes `InlineWrapper`, discriminated on `kind`, with `type: "inlineWrapper"`. The old name admitted only one kind of transparent wrapper; a smart tag and a custom-XML wrapper are the same shape and become added members rather than new types every exhaustive switch has to learn. Only `bidi` exists today: nothing parses the other two yet.

`AUTOSAVE_FORMAT_VERSION` moves to 3, because the codec serialises the model and an envelope written under 2 holds paragraph content under the old discriminator. A stored collaboration snapshot is unaffected and the attr-schema version does not move: the new mark attr defaults to `null`, which is what every existing snapshot means.

The save leg is unchanged. An edited wrapper span still loses its wrapper, because `fromProseDoc` rebuilds the wrapper from the source paragraph rather than from the mark.
