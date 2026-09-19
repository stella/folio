---
"@stll/folio-core": patch
---

Comparing a document with itself reports no change in three cases where it used to refuse or invent one. An inline atom the comparison cannot detach from its package — an image with no embedded media, a drawing that is a chart rather than a picture — is now compared as part of the block's topology instead of making the whole story unalignable. A package with no `word/styles.xml` compares against another that has none: two packages with no style definitions share one formatting context, so there is nothing to isolate and nothing to read a definition from. And two structurally identical tables no longer cancel each other's pairing: each shares every signature with the other by construction, and evidence that also exists locally no longer counts as evidence that the counterpart lies elsewhere.
