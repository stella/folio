---
"@stll/folio-core": patch
---

Render complex-script text in the font the document asks for.

Word resolves a run's font per character across three slots: `w:eastAsia` for CJK, `w:cs` for Arabic, Hebrew, Indic and South-East Asian text, and `w:ascii`/`w:hAnsi` for everything else. folio honoured the first and third but never the second: `w:cs` was parsed and round-tripped yet never reached layout, so a document written the standard way (`w:ascii="Calibri"` with `w:cs="Traditional Arabic"`) measured and painted its Arabic in Calibri.

The complex-script slot now flows through the bridge, the measurer and the painter on the same path the East-Asian slot already used, so the two stay segmented identically and line wrapping keeps matching what is drawn. Script segmentation now reports which of the three slots a character selects rather than a CJK yes/no.
