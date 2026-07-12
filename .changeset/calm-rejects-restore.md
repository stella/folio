---
"@stll/folio-core": patch
---

Property-change tracked revisions now resolve fully. Rejecting a `w:pPrChange` restores the stored old paragraph properties wholesale within CT_PPrBase scope (properties the change added are cleared, out-of-scope attrs like the inline `sectPr` still preserved). `w:sectPrChange` and table property changes (`w:tblPrChange`/`w:trPrChange`/`w:tcPrChange`) — previously display-only — now accept and reject: accept keeps live values and clears the record, reject restores the stored previous properties (section rejects keep live header/footer references, which CT_SectPrBase cannot carry). Table property-change records also survive the ProseMirror round-trip instead of being dropped on save, and `acceptAll()`/`rejectAll()` counts include section and table property changes.
