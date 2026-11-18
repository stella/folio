---
"@stll/folio-core": patch
---

Move a pending deleted break onto a paragraph inserted after its words as soon as the insertion lands, so later operations of the batch read those words as after the insertion alone; and let a kept join break go when a later batch deletes the words joined into it, as one batch deleting both does.
