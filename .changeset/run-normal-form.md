---
"@stll/docx-core": patch
---

Every operation leaves the paragraphs it touched in their run normal form: consecutive runs that can merge, at the paragraph's top level and inside hyperlinks, become one run, as the parser consolidates them, so the same text, formatting and ids have one run structure whichever order the edits arrived in. The joins are part of the edit, and its inverse splits them first, so undo restores the structure exactly. Splits, inline, block, container and story replacements, and edits of a paragraph that was not in normal form keep runs as they are; untouched paragraphs are never changed. `runSeams` lists the seams a paragraph's normal form joins.
