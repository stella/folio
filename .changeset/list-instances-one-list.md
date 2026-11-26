---
"@stll/folio-core": patch
---

List numbering counts the instances of one list definition as one list, as the reference consumer does. A paragraph goes on from the last paragraph of its definition, whichever instance that one named, also across other paragraphs in between. An instance's start override sets the number once, on the first paragraph that uses the instance on that level. Before, each instance kept its own count except in a few adjacent cases, so a list that returned to an earlier instance after a restarted one skipped a number (`a)`, `b)`, `a)`, `c)` for what is printed `a)`, `b)`, `a)`, `b)`).
