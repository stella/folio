---
"@stll/folio-core": patch
---

Merge two adjacent runs only when their whole record agrees, and keep the record the merge inherits.

`consolidateRuns` compared typed formatting alone and rebuilt the merged run from `type`, `formatting` and `content`, so a run's attribute remainder and its property set's verbatim sink were dropped at parse time, before anything downstream could see them. Two runs written in different editing sessions became one run in neither, and where the sink differed the survivor's captured bytes were applied to the other run's text as well: a `w:webHidden` on one run hid the text beside it.

One predicate, `runsMergeable`, now answers the question for every consolidation site: formatting, attribute remainder and `w:rPr` sink must all be equal. The remainder compares as a set, because attribute order in XML says nothing; the sink compares as a sequence, because its order is what puts the markup back between the same modelled siblings. The merged run is the survivor spread whole, so no field can be dropped by a field list that forgot it.

Both field lists the predicate reads are now total over their model type, which turned up a second field the merge had been crossing: `w:noProof`.

More of a document's runs survive a parse as a result, by roughly a third on files that carry run-level session ids.
