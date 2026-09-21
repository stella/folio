---
"@stll/folio-core": patch
---

Stop a comparison from recording a run-property change that changes nothing.

Replacing text clears a highlight or `w:shd` under it, because text typed over a highlighted placeholder is new text and the marker that said "fill this in" should not survive into the finished document. A comparison is not authoring: it holds the revised document's own run properties and writes them itself. Clearing them first recorded a `w:rPrChange` that the provenance pass then took straight back, so every carrier of an edited paragraph in a highlighted cell reached the reader as a revision whose before and after were identical, and the run the replacement deleted kept a claim that its background had gone.

`replacementBackground` names the two behaviours on the apply path, defaulting to `clear`; `compareDocx` asks for `keep`.
