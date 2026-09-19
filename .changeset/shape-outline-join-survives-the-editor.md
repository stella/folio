---
"@stll/folio-core": patch
---

Carry a shape outline's `a:ln@join` through the ProseMirror model. The parser
read it and the serializer wrote it, but the shape node had nowhere to put it,
so a mitred or bevelled outline came back rounded after any edit.
