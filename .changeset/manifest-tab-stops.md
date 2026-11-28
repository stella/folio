---
"@stll/folio-core": minor
---

The block manifest gives a paragraph's tab stops in effect (`tabStops`: position, alignment and leader, from its style chain and its own `w:tabs`, cleared stops left out). A template plan item's `tabStops` may name a stop's alignment and leader beside a bare left position, and `tabStopsOnly` clears every other stop the paragraph's style sets, so a line laid out by tabs (a signature block) keeps the source's stops. A plan may give a paragraph up to sixteen stops.
