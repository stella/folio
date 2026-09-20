---
"@stll/folio-core": patch
---

Read a style's run properties with the same reader as a run's, and resolve a property stated twice to its last statement.

`styleParser.ts` kept a private copy of `parseRunProperties`, and the copies had drifted. The style copy never read `w:noProof`, wrote an empty `fontFamily` and an empty `color` for a `w:rFonts` or a `w:color` that stated nothing, kept none of the `w:rPr` children folio does not model, and — the difference that changed a value — took the *last* of two statements of a toggle while the run copy took the first. It also disagreed with itself: `w:rtl`, `w:cs` and `w:dstrike` took the first statement while `w:b` and `w:strike` took the last.

`EG_RPrBase` is an `xsd:choice` referenced `maxOccurs="unbounded"`, so a repeat is valid markup rather than a malformed file: 44 of the 5299 packages in the public corpus hold one. The last statement wins, and the statements it beat are not written back, so the saved element states each property once and a consumer that takes the first and one that takes the last read the same value from it. The evidence is recorded as `repeated-run-property-resolves-last`.

A style's `w:rPr` now also keeps what folio models nothing for, the way a run's already did.
