---
"@stll/folio-core": minor
"@stll/folio-react": patch
"@stll/folio-vue": minor
---

Use canonical review to predict editor lock outcomes. The deep ProseMirror accept/reject commands and their lock probes are now test-only. Pass the canonical editor to Vue's `useTrackedChanges` to include lock verdicts in sidebar entries.
