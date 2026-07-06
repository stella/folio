---
"@stll/folio-vue": minor
---

Wire the tractable remaining chrome/callback props on the Vue `DocxEditor`,
moving them from deferred to paired with the React adapter: `showRuler` +
`rulerUnit` (renders the horizontal and vertical rulers, gated on `!readOnly`,
with live margin / indent / tab-stop editing), `autoOpenReviewSidebar`,
`initialScrollTop` + `onScrollTopChange`, `onFontsLoaded`,
`onCompatibilityChange`, `onSelectionTextChange`, `placeholder` /
`loadingIndicator` (nodes or `#placeholder` / `#loading-indicator` slots), and
`customContextMenuItems` + `onCustomContextAction` (host-injected right-click
entries that lead the menu, with the selection range captured at open time).
Genuinely-hard or React-inert props (header/footer editing, template overlay,
anonymization overlays, host FontFace registration, keep-previous-document-
while-loading, margin guides, and the clipboard / review-control / print-button
flags React itself leaves unwired) stay deferred with updated reasons.
