---
"@stll/docx-core": minor
"@stll/folio-core": minor
---

Keep `w:bdo` and `w:dir`, the Unicode bidirectional controls. `w:dir` is an embedding and `w:bdo` an override, and folio discarded both on save — in a right-to-left document that is the difference between a readable line and a scrambled one, because an override is what makes a Latin word inside it read backwards. They are now a `BidiWrapper` member of `ParagraphContent`, a transparent inline container that nests, holds anything paragraph content holds, and that every paragraph walk reads straight through. The editor projection flattens the wrapper for now, so a document edited in the editor still loses the direction; the save path keeps it.
