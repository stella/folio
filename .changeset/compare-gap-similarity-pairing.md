---
"@stll/folio-core": patch
---

Pair blocks by similarity inside a comparison gap that gained or lost blocks. Between exact anchors, a gap whose two sides differ in length used to pair its blocks by position, so an inserted block was read as a rewrite of the block it pushed down, and each later block as a rewrite of its predecessor's wording. Such a gap now keeps the order-preserving pairs with the greatest summed word similarity (multiset Dice, each pair at least 0.5), reading the rest as inserted or deleted; a gap of equal sides still pairs by position. A split or merged paragraph is recognised whichever half pairs with the whole.
