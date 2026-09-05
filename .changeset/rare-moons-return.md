---
"@stll/folio-core": patch
---

`compareDocx` returns the base package unchanged when it finds no difference,
instead of re-serializing it. A 2,200-block comparison of two identical
documents spent a second rewriting bytes nobody edited. A base that arrived
carrying its own tracked changes is still serialized, because the compared base
is its accepted view rather than the package as stored.
