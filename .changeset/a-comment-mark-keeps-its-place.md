---
"@stll/folio-core": minor
---

Carry `w:commentReference` through the editor as an inline node of its own, so
a comment's visible mark keeps the place it was authored in. The mark-only
projection recorded no position for it and the serializer guessed one after
every range end: two comments closing together came back interleaved rather
than grouped, and a comment spanning three paragraphs came back marked three
times. The serializer now writes what the model says, the schema gains a
zero-width `commentReference` node whose integrity is repaired on every
transaction, and a model that arrives with a range and no reference is
completed once for the whole story rather than per paragraph.
