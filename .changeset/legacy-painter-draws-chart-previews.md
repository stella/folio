---
"@stll/folio-core": patch
---

The default painter now draws chart and diagram previews, which it used to leave blank, and a chart with no drawable data keeps the height of its frame on both painters instead of collapsing and moving the text after it. Chart text (titles, axis labels and legends) is still not drawn.
