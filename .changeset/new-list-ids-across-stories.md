---
"@stll/folio-core": patch
---

A list a headless document operation starts in one story (the body, a header, a footer or a note) takes a list id no other story of the package uses. Lists started in two stories no longer share an id until a save moves one of them, so the list id a reader reports stays the same across save and reopen.
