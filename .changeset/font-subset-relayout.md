---
"@stll/folio-core": patch
"@stll/folio-react": patch
"@stll/folio-vue": patch
---

Czech, Polish, Greek and Cyrillic text in a bundled font is measured in that font, not a fallback: the initial layout loads every `unicode-range` subset the document's characters need, and a face that loads after a layout (a subset first needed by new text, or one the initial wait timed out on) re-lays out the document. Measurement caches and incremental measures are bound to the font set they were taken in, so a fallback measurement is never reused once the real face has loaded.
