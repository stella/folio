---
"@stll/docx-core": minor
"@stll/folio-core": minor
---

Keep every `w:rPr` child folio does not model, where it stood, for a run, a paragraph mark, a style and a tracked property change alike.

A run property set was walked by a reader per property with no branch for the rest. `w:bdr`, `w:fitText`, `w:eastAsianLayout`, `w:snapToGrid`, `w:webHidden`, `w:specVanish` and `w:oMath` had no model at all; `w:rFonts`, `w:u`, `w:lang`, `w:w` and `w:sz` were read and dropped whenever the reader took no typed value from them; and the paragraph mark's whole `w:rPrChange`, along with everything inside it, went with them. A save that rewrote the element — which is every save after an edit — lost all of it.

`EG_RPrBase` now goes through the shared child dispatcher, with a handler map the compiler makes total over the children the schema declares. A handler answers with what it took, so a property the reader turned into no typed value keeps its bytes: a name-keyed map can state the names folio has never heard of, not the values a reader refuses. `TextFormatting` gains `preserved`, and because it is a sequence the sink records each capture's schema ordinal rather than a count of modelled siblings.

The four owners of a run property set — a run, the paragraph mark inside `w:pPr`, and the snapshot inside either one's `w:rPrChange` — share that map and differ only in which children a sibling record has already claimed, which the call site names. One writer serves all of them plus a style and a numbering level, and it orders its children from the generated declared-child list rather than from the order of its own statements: folio wrote `w:vanish` before `w:noProof` while the schema declares the reverse, which a validating consumer refuses.

Captured bytes belong to the element that was parsed and to no other, so style resolution and formatting merges drop them rather than inheriting them onto every run below.
