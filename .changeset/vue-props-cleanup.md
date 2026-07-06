---
"@stll/folio-vue": minor
---

Squeeze the Vue adapter's remaining prop deferrals to 58/64 paired. Wire three
real props: `toolbarExtra` (the `VNodeChild` prop now renders into the toolbar's
`#toolbar-extra` outlet, with a host-provided slot taking precedence),
`preserveDocumentWhileLoading` (suppresses the loading interstitial on a document
swap once a document has painted, so the previous pages stay visible until the
new layout replaces them), and `fonts` (ports React's host `FontFace`
registration + font-ready re-layout via a DOM-only `utils/hostFonts` module).

Reclassify six no-op-in-both props to paired after verifying React declares but
never consumes them: `showMarginGuides`, `marginGuideColor`, `showPrintButton`,
`onCopy`, `onCut`, `onPaste`. `showReviewControls` stays deferred (a real gap:
React renders review controls the Vue toolbar lacks).
