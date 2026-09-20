---
"@stll/folio-core": patch
---

Report an edited table cell as an edit inside its row, in a document whose
producer wrote no `w14:paraId`. The row alignment paired a container only when
its block ids were identical on both sides and their stability had changed from
positional to stable — the shape a save left behind when it stamped folio's
minted ids into the package. That made the comparison depend on a side effect
of the save rather than on the two documents in front of it: with minted ids no
longer persisted, both sides read positional, the pairing never fired, and one
edited cell came back as a deleted row plus an inserted row. Container identity
now goes through `alignParagraphOrdinals`, the same owner the selective save
asks which paragraph is which, so the two cannot answer differently.
