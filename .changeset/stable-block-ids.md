---
"@stll/folio-core": patch
"@stll/folio-cli": patch
---

Keep block ids stable across changes to a file whose paragraphs carry no `w14:paraId`: `ensureParaIds` reports the ids it minted, and the CLI opens every file with paragraph ids, so the first change stores the ids a read reported.
