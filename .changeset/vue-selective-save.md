---
"@stll/folio-vue": minor
---

Wire selective-save feature flags and the tripwire callback into the Vue
`DocxEditor` save path, matching the React adapter. `featureFlags` is threaded
into `useDocxEditor`'s `save()`: when `selectiveSave` is on it runs
`attemptSelectiveSave` (patching only changed paragraphs, with the shared
`ParagraphChangeTracker` inputs) and falls back to a full repack; the ref
`save({ selective })` option now gates that path. When `selectiveSaveTripwire`
is on it also computes the full repack, runs `compareSelectiveVsFull`, and fires
`onSelectiveSaveTripwire(result)` for observability. The comparison never blocks
or poisons the save. Both props move from deferred to paired in the cross-adapter
parity contract.
