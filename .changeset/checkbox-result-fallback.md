---
"@stll/folio-core": patch
"@stll/docx-core": patch
---

Stop a legacy `FORMCHECKBOX` field's synthesized display glyph from being written into `w:sdtContent`/the field result on save when the field carries no cached result of its own. The parser still models the glyph so the editor can paint it, flagged as a display-only fallback (`ComplexField.fieldResultIsFallback`) the serializer now honours by leaving a resultless field's result empty, matching the source. `w:sdt`/`w14:checkbox` content controls already round-tripped their authored `w:sdtContent` correctly and are unaffected.
