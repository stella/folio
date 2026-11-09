---
"@stll/folio-core": patch
---

Typing in a table cell repaints only the changed line: the painter replaced the whole table for any change inside it, and the browser laid out and painted every cell again.
