---
"@stll/folio-core": patch
"@stll/folio-react": patch
"@stll/folio-vue": patch
---

Opening the document outline in the Vue editor gives it a track of its own beside the page, as in the React editor, instead of a drawer over the page. Both editors take the opened outline's layout and when it closes from the same shared panel rules (`requestedOutline`, `overlayAfterLayout`).
