---
"@stll/folio-vue": minor
---

Port the anonymization + template-directive decoration overlays to the Vue
adapter. Both project the existing core plugins' decoration ranges onto
page-relative rects via `projectRangesToRects` (DOM-rect primary path, layout
fallback) and paint them over the pages. `AnonymizationRectsOverlay.vue`
highlights redaction matches, forwards clicks, and scrolls the selected term
into view; `TemplateDirectivesOverlay.vue` tints `{{…}}` markers and draws the
per-page block gutter rails. The anonymization, template-directive, slash-menu,
and template-preview plugins are now registered in the hidden-editor pipeline
(template plugins gated on `showTemplateDirectives`, reconfigured live on
toggle). Moves `onAnonymizationMatchesChange`, `onAnonymizationTermClick`,
`selectedAnonymizationCanonical`, `anonymizationSelectionSeq`,
`showTemplateDirectives`, `onSlashMenuChange`, and `onSlashMenuKeyAction` from
deferred to paired in the cross-adapter parity contract.
