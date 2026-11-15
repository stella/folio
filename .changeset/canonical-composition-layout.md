---
"@stll/folio-core": patch
"@stll/folio-react": patch
"@stll/folio-vue": patch
---

A canonical session lays out an IME composition's text as it grows, so the painted text and caret follow it, and reports the composition to the host only once it commits. The editor draws a run's text, a composition included, with the faces the page paints it with, East Asian slot and embedded fonts included, so the caret the browser shows stays beside the painted one.
