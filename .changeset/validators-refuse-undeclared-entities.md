---
"@stll/docx-core": minor
"@stll/folio-core": minor
"@stll/folio-cli": patch
---

A part holding a reference XML refuses is not opened and does not validate: an entity no declaration defines, or a character reference to a character XML 1.0 does not allow (`&#0;`, `&#x1;`, `&#xFFFE;`, a surrogate, past U+10FFFF), with digits that are not a number (`&#12ab;`), or with a hexadecimal marker in capitals (`&#X41;`). Both engines refuse it with `MalformedXmlReferenceError` (part, offset, reference and kind), which replaces `UndeclaredXmlEntityError`; `malformedXmlReferenceError` replaces `undeclaredXmlEntityError`. The legacy engine no longer reads such a reference as nothing. The package validators and the CLI's save checks refuse the same parts. `XmlReadOptions.undeclaredEntities` and `XmlUndeclaredEntities` are removed; no read accepts these references.
