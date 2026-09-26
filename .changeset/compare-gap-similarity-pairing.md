---
"@stll/folio-core": patch
---

Pair blocks by similarity inside every comparison gap. Between exact anchors, a gap used to pair its blocks by position, so an inserted block was read as a rewrite of the block it pushed down, and each later block as a rewrite of its predecessor's wording; a gap of equal sides hid the same misreading whenever an insertion sat beside a deletion. A gap now keeps the order-preserving pairs with the greatest summed word similarity (multiset Dice, each pair at least 0.5), reading the rest as inserted or deleted; blocks between those pairs still pair by position when their counts match and none of them resembles any block across the gap. A split or merged paragraph is recognised whichever half pairs with the whole.
