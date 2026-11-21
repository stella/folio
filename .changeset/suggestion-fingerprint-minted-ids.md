---
"@stll/folio-core": patch
---

Host-persisted pending suggestions restage after a save on documents whose paragraphs name no `w14:paraId`: the source fingerprint no longer counts whether a paragraph id was minted at open, which a save that writes the id changes.
