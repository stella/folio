---
"@stll/folio-cli": patch
---

Scale `folio serve`'s preview page down to fit a browser window narrower than the document, instead of clipping the left margin and scrolling horizontally. Fix the herdr plugin's `npx` fallback so it no longer fails when the plugin runs from inside a folio checkout, and resolve a relative clicked `.docx` path against the pane it was clicked in.
