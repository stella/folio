---
"@stll/docx-core": minor
"@stll/folio-cli": patch
---

The package validators refuse an XML part that names an entity no declaration defines, as both engines do when opening it. `XmlReadOptions.undeclaredEntities` and `XmlUndeclaredEntities` are replaced by `characterReferences: "xml" | "syntactic"`, which only chooses the character references a well-formed read admits; no read accepts an undeclared entity any more.
