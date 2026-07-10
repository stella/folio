---
"@stll/folio-react": patch
---

Translate the display-mode select trigger label (it was hardcoded English while the popup was localized) and keep the items array referentially stable per locale instead of rebuilding it every render.
