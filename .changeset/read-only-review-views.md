---
"@stll/folio-core": minor
"@stll/folio-react": minor
"@stll/folio-vue": minor
---

Add controlled review display modes and independent comment highlight visibility with search and layout following the displayed document.

Projected review views refuse editing and report a typed read-only reason. `getPageTextFromLayout` now takes one options object (`{ layout, doc, page, projection? }`) instead of three positional arguments.
