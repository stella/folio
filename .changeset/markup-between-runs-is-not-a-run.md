---
"@stll/folio-core": patch
---

Stop counting markup between runs as a run-formatting carrier.

A preserved capture is a `w:r` child or a `w:p` child, and only the first serializes inside a run and owns a `w:rPr`. Both arrive as the same editor node, and the map from node type to run-formatting carrier could only name the type, so the paragraph-level capture — `w:proofErr` above all, which Word writes between the runs of any sentence its grammar checker flags — was classified as a run. Formatting marks put on it were dropped on the way back out, and a comparison was asked to line the base document's proofing annotations up with the revised document's own: no redline can, so the round-trip check refused redlines whose content was right, reporting `inline-formatting`. The classification now reads the level the capture came from, which is the level the save path already branches on.
