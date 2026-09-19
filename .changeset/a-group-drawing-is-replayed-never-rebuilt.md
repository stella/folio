---
"@stll/folio-core": patch
---

Replay a preview-only drawing instead of rebuilding it. A `wpg:wgp` group is
rendered to an SVG preview at parse time, and the model holds the render, not
the group; once the preview's fingerprint no longer matched — after any edit,
and on every rebuild path — the serializer regenerated DrawingML from the
render, dropping the group's children, their relationships and the preview's
own filename. Both classified raw-XML modes are now written back whatever the
model says. A stale fingerprint still classifies the drawing `opaque`, which is
how the lost edit is reported; it no longer licenses a replacement.
