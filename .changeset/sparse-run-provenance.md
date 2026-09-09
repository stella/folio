---
"@stll/folio-core": patch
---

Preserve direct run formatting, character-style inheritance, and formatted inline carriers (tabs, line breaks, symbols, and fields) across editor and DOCX round trips while keeping common run state compact and reviewable. Refuse a second independently owned run-property revision before mutation, and commit comment and revision ids only with the operation that writes them.
