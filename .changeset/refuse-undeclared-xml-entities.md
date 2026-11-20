---
"@stll/docx-core": patch
"@stll/folio-core": patch
"@stll/folio-cli": patch
---

The legacy engine no longer opens a package whose XML names an undeclared entity (`&nbsp;`, `&copy;` or any name but XML's five built-ins); it used to read common HTML names as their characters and keep other names as written. Both engines now refuse such a package with `UndeclaredXmlEntityError`, which names the part, the reference's offset and the entity. Character references are read as before. The CLI's save check refuses such a part as not well-formed.
