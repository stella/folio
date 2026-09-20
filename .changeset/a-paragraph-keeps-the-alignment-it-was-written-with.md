---
"@stll/docx-core": minor
"@stll/folio-core": minor
---

Generate `ParagraphAlignment` from `ST_Jc` instead of spelling it by hand. The
union omitted `start`, `end` and `numTab`, so a paragraph written with one
parsed without an alignment and saved without a `w:jc`. `start` and `end` are
direction-aware members, not spellings of `left` and `right`: the layout and
the CSS projection resolve them against the paragraph's direction.
