---
"@stll/folio-core": patch
---

Keep a shifted row the two sides hold identically.

A sole shifted pair that also strands containers on both sides was split back into a deletion and an insertion whenever the container count was unchanged, so a table that lost two rows and gained two reported every row deleted and every row new, the surviving one included.

That rule is about a mapping inferred from similarity, and a pair whose content digests are equal and whose blocks carry the same ids in the same order is not one: there is nothing left to infer. Both halves are required, since equal content alone is two boilerplate rows reading alike, and equal ids alone are positional ids agreeing after a reorder with no content behind them.
