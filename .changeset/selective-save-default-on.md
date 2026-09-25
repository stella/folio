---
"@stll/folio-core": minor
"@stll/folio-react": minor
"@stll/folio-vue": minor
---

Selective save is now the default on the save path instead of an opt-in flag. Every save still falls back to a full repack automatically whenever the patch-safety checks refuse, so a save always succeeds. Pass `selectiveSave={false}` to restore full repacking.
