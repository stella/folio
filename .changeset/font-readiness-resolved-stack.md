---
"@stll/folio-core": patch
---

Wait for the fonts the renderer actually uses, not just the names the document wrote.

`resolveFontFamily` turns an authored family into a CSS stack that appends folio's bundled metric-compatible substitutes and a script fallback, so an authored `Arial` run paints its Arabic in the bundled Arabic face. The font-readiness gate collected only authored names, so it released the first layout before those faces had loaded; measurement taken against the pre-load fallback then disagreed with what was ultimately painted, by as much as a third of a line's width.

The gate now expands each family through the resolver and waits for every concrete face in the stack. That also removes a hand-kept substitute table which duplicated the resolver's own mapping and was free to drift from it.
