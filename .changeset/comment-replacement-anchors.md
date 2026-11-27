---
"@stll/folio-core": minor
---

Typing over, replacing or deleting text that holds comment anchors now applies in every mode, including suggestions and selections across paragraphs. Anchors inside the replaced text stay, in order, at the start of the new text; a ranged comment whose words all went is removed with them, and a point comment keeps its thread. Tracked, the deleted text keeps its anchors, so rejecting restores them and accepting matches the direct edit. The `comment-replacement-alignment` capability gap is gone.
