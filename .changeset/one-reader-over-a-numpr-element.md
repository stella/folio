---
"@stll/docx-core": minor
"@stll/folio-core": minor
---

Give `w:numPr` one union, one cascade fold and one reader.

`ParagraphNumberingOverride` is what a tier states — `none` (the reserved `w:numId 0`), `reference` (an id and an optional level) or `levelOnly` — and `ResolvedParagraphNumbering` is what the cascade leaves. Three arms, not two: `w:numId` and `w:ilvl` inherit independently (ECMA-376 17.3.1.19), so a tier that states only the level keeps the id it inherits, and that shape is what Word writes whenever a styled list paragraph is demoted.

`mergeParagraphNumbering` is that inheritance written once. It replaces the paragraph parser's object spread, and it is closed under itself and associative over the three cascade tiers, so a third tier needs no special case.

`paragraphNumberingFromSlots` is the one mapping from the element's two slots onto an arm, and `readParagraphNumbering` reads the element. Both are exported from `@stll/folio-core/docx` alongside `NO_NUMBERING_NUM_ID` and `isNumberingReference`, which now live in `@stll/docx-core` where the model does. Three duplicate spellings of the reserved id are retired: the hand-inlined copy in `docx-core`'s validator, the bare literal in the operation reader, and the relational form in the AI snapshot, which was the one spelling that read a malformed package's negative id as "not numbered" while every other spelling read it as a dangling reference.
