---
"@stll/folio-core": patch
---

Stop inventing a drawing's `a:graphicFrameLocks`. Rebuilding a picture whose model held no lock record wrote `noChangeAspect="1"`, because one `undefined` meant both "the author wrote no frame properties" and "folio created this picture": at serialization the two are indistinguishable, so an edited document gained a lock its source never carried. The default now belongs to the insert that creates a picture, and the serializer writes the element only when the model holds locks.
