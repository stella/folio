---
"@stll/folio-core": patch
---

In suggesting mode a list toggle on a paragraph that already carries a tracked paragraph-property change applies as a further tracked change (one record whose reject restores the original properties) instead of doing nothing, and list autoformat converts typed markers as a tracked change; Backspace right after the conversion puts the marker back.
