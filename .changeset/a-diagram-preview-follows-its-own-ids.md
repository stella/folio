---
"@stll/folio-core": patch
---

Resolve a SmartArt preview's drawing cache through the diagram's own `r:dm` id and the data part's `dsp:dataModelExt`, rather than by scanning the relationship map for the drawing type. The scan could serve one diagram another's drawing, and refused outright on a second match, so a document with two diagrams got a preview for neither.
