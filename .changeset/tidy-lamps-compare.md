---
"@stll/folio-core": minor
"@stll/folio-agents": minor
---

Version comparison upgrades and a redline generator. `compareDocxVersions` now detects relocated blocks (`movedFrom`/`movedTo` pairs sharing a `moveGroupId`) instead of reporting them as unrelated delete + insert, and reports text-equal blocks whose run formatting differs as `formatChanged` with the changed property names. New `generateRedlineDocx(base, revised)` produces a third `.docx` whose base → revised differences are recorded as real Word tracked changes (`w:ins`/`w:del`), reusing the comparer alignment and the headless tracked-changes apply path. `formatVersionDiffForLLM` renders the new change types.
