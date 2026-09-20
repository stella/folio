---
"@stll/folio-core": patch
---

Run the collaboration attr-schema migration on every path that reads a stored fragment, not only in the offline sweep. A snapshot written under an older attr schema reached `initProseMirrorDoc` unmigrated, so a value a step rewrites was read in its old shape as the new one; and the marker was stamped only when Folio happened to rewrite the whole fragment, so an editor could write this build's attrs into a fragment still marked older, which an older build would then read and drop unnoticed. `applyAttrSchemaMigrations` owns both, and the editor, the server materialization and `migrateFolioYjsSnapshot` all go through it.

A rotate from the editor states the turn the drawing is at. `%` keeps the sign of its left operand in JavaScript, so rotating a drawing whose authored `rot` was negative answered with a negative rotation.
