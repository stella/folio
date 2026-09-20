---
"@stll/docx-core": minor
"@stll/folio-core": minor
---

Keep every `w:pPr` child folio does not model, in the place the schema gives it, through a rebuild as well as through a replay.

`CT_PPrBase` declares thirty-three properties and folio models fourteen. The rest reached disk only while the whole `w:pPr` was replayed as bytes, so the first edit to any paragraph property — an alignment command, a spacing change, a style applied — rebuilt the element without them: a table's `w:cnfStyle`, an East Asian document's `w:wordWrap` and `w:autoSpace*`, a frame's `w:mirrorIndents`, a vertical text box's `w:textDirection`. A style's property set was read by a second, narrower copy of the same `if`-chain, which knew neither `w:framePr` nor the Strict `w:ind` spellings.

One reader now dispatches the set through the shared child dispatcher, so the compiler makes the handler map total over the declared children and each one carries a decision: modelled, kept as bytes, or named as another reader's. `w:rPr`, `w:sectPr` and `w:pPrChange` are the three with other owners. A handler that takes no typed value hands the child back instead of dropping it, which is what `<w:spacing/>` and a `w:jc` outside the reader's enumeration used to do.

`ParagraphFormatting` gains `preserved`, the sink, recording each capture at its **schema ordinal** rather than at a count of modelled siblings: the count is a mirror of whichever properties folio models today, and it moves under the capture the moment one more of them is modelled. The order the four writers emit comes from a generated table in `@stll/docx-core/schema`, and one writer serves all four — a paragraph, a style, a numbering level, and the `CT_PPrBase` snapshot inside `w:pPrChange`.

The cascade drops the sink rather than inheriting it: captured bytes belong to the element they were read from, and writing a style's back as direct formatting would outrank the tier they came from.

A `w:pPrChange` whose original states nothing is no longer discarded for being empty. It records that the paragraph carried no direct formatting before the reviewer's edit, which is what rejecting the revision restores.
