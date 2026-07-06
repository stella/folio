---
"@stll/folio-vue": minor
---

Wire the review controls in the Vue toolbar, gated on `showReviewControls`
(default true), matching the React adapter. Adds a track-changes toggle
(flips editing<->suggesting through the existing `setSuggestionMode` path) and a
markup display-mode selector (All Markup / Simple / No Markup / Original) that
applies the same `folio-root--<mode>` root class and tracked-change display CSS
as React. `showReviewControls` moves from deferred to paired in the
cross-adapter parity contract.
