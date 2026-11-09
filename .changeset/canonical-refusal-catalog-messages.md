---
"@stll/folio-core": minor
"@stll/folio-react": patch
"@stll/folio-vue": patch
---

A change the experimental canonical session cannot make yet now shows a translated message from the folio catalog ("This change can't be made here yet.") instead of the session's internal description. `canonicalRefusalNotice(ground, translate)` no longer takes the session's message: every ground reads a catalog key, and the new `editor.changeUnavailable` message is in every shipped locale. The session's message stays on the error passed to `onError`.
