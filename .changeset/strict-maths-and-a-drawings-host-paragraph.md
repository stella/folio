---
"@stll/folio-core": patch
---

Read maths in an ISO Strict package as an equation, and keep an anchored drawing's host paragraph attributes.

The child dispatcher looks a namespace-keyed disposition up by the URI the element carries, and the dispositions are written in Transitional. Strict spells the maths namespace `purl.oclc.org/ooxml/officeDocument/math`, so a Strict document's `m:oMath` missed the lookup and went to the verbatim sink: the bytes survived and the equation stopped being one, with nothing left to render, edit or read text from. Both the dispatcher and the reader behind the disposition now resolve the namespace through the generated Strict/Transitional pair table, so the rule holds for every namespace-keyed disposition rather than for maths alone.

Word writes a floating shape into a paragraph of its own. folio lifts the shape out as a block node and drops that paragraph from the editor projection, so the attributes the `w:p` carried and the model has no field for, `w:rsidR` and its family, had no carrier and were gone on the way back. The text box node stands in for the host paragraph and now carries the host's remainder, which the save leg puts back on the paragraph it rebuilds; only the first node of a group takes it, because one paragraph is rebuilt for the group.
