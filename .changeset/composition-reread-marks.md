---
"@stll/folio-core": patch
---

In the canonical session, a composition that continues after the editor redraws its text, for example beside pasted text with a font set, keeps going instead of being refused: reading the text back may now lose mark attributes that the rendering cannot carry.
