---
"@stll/folio-core": patch
---

`compareDocx` is faster on large documents, up to about twice as fast on a heavily edited one: checking run formatting no longer re-walks the document once per run, and a target without section breaks in the body no longer costs an extra accept-all pass.
