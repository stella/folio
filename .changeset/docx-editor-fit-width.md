---
"@stll/folio-react": minor
---

Add `initialZoom="fit-width"` to `DocxEditor`. It sizes the page to the editor's width and keeps it fit as the editor resizes (never enlarging past 100%), so a document embedded in a container narrower than its page no longer overflows and clips on the right. A manual zoom (toolbar or `ref.setZoom`) afterwards overrides the fit. The fit is computed from the laid-out page geometry, so the scaled page never has to be measured in the DOM.
