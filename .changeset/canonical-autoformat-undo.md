---
"@stll/folio-core": patch
---

Undoing a list or heading autoformat in a canonical session gives back exactly what was typed, the trigger space included: the typed text is committed first and the rule is a history step of its own, so Backspace straight after the rule or undo restores "- " rather than "-".
