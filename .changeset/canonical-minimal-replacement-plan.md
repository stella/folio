---
"@stll/folio-core": patch
---

The canonical session plans text replacements with the shared minimal planner: a direct replacement writes only the characters it changes, a redline keeps every field the direct edit keeps, a change may span a repeated line break, the background clears from every highlighted stretch a change touches, and a range cutting into a field result is skipped as `unsupportedBlock`.
