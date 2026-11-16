---
"@stll/folio-core": patch
"@stll/folio-react": patch
"@stll/folio-vue": patch
---

A direct `mergeBlockWithNext`, and a direct `deleteBlock` of a container's final paragraph that a paragraph precedes, leave the paragraph whose mark remains, the later one, as accepting the same edit tracked does: both engines keep its block id. The applied operation reports `paragraphIdentity`: a `joined` outcome names the surviving and retired ids; a `retained` outcome names an identity that remains without retiring another paragraph. Receipts report surviving ids as updated and only retired ids as deleted.
