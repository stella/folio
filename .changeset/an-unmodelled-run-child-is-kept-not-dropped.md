---
"@stll/docx-core": minor
"@stll/folio-core": minor
---

Keep every run child folio does not model instead of letting it fall off the end of the run-content switch. `RunContent` gains a `preservedXml` member holding the captured markup at its source position, plus the visible text it contributes, so `w:ruby`, `w:contentPart`, `w:pgNum`, `w:annotationRef`, the note markers and any foreign or future element survive a save and read as text.

The keep rule now asks the model rather than the source element. The two disagreeing was a two-save oscillation rather than a loss: the first save wrote a run whose payload the model never held, the next parse dropped that run, and the second save differed from the first.
