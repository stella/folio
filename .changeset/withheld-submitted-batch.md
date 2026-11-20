---
"@stll/docx-core": patch
---

The optimistic sequencing client withholds a submitted batch that no longer rebases until the authority answers it: a rejection drops it with a notice, and a sequenced one comes back as a document change instead of being reported dropped and then applied unseen.
