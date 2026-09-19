---
"@stll/docx-core": minor
"@stll/folio-core": patch
---

Keep every `w:tblLook` the author wrote. `CT_TblLook` states which conditional formats a table takes from its table style twice: as the legacy `w:val` bitmask, which `TableLook` had no field for, and as six `ST_OnOff` attributes, which the serializer wrote only when true. A rebuild therefore turned `w:val="04A0" w:firstRow="1" w:lastRow="0" w:noHBand="0"` into `w:firstRow="1"`, and the two are different documents: an absent flag falls back to `w:val`'s bit, an explicit `0` overrides it. `TableLook` gains `val` and each flag is now tri-state. Two readers also disagreed about precedence — the table one OR-ed `w:val`'s bits over an explicit `0`, so a table that switched its header row off got one anyway; `styleParser` now calls the table parser, and `resolveTableLook` is the single place a flag resolves to an answer.
