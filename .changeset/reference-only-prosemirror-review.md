---
"@stll/folio-core": minor
"@stll/folio-react": patch
"@stll/folio-vue": minor
---

Use canonical review to predict editor lock outcomes. Remove imports of the legacy accept/reject commands and lock probes from `@stll/folio-core/prosemirror/commands/comments` and `@stll/folio-core/prosemirror/commands/index`. Use the editor ref's `acceptAIEditOperation`, `rejectAIEditOperation`, `acceptSuggestion` or `rejectSuggestion` to resolve changes. Existing two-argument calls to Vue's `useTrackedChanges` now omit lock verdicts. Pass the canonical editor as the third argument to keep them.
