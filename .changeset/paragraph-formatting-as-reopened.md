---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

Paragraph properties an edit sets read in the live model as the saved file reopens: the spacing sides the paragraph states are marked as the parser marks them, and a hanging indent is held as the parser reads it, a negative first line flagged hanging. `paragraphSpacingExplicit` is the one rule for the spacing sides, used by the parser and by every edit that finishes a paragraph's formatting.
