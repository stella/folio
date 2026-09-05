---
"@stll/folio-core": patch
---

`compareDocx` compares the accepted view of both documents. An input that
already carried tracked changes previously produced a package with two
redlines layered on one another, where rejecting everything landed on a
document neither side wrote; the base's own revisions are now resolved first,
so the comparison is the only redline in the result and the round trip is
exact.
