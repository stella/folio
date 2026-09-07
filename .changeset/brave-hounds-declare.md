---
"@stll/folio-core": patch
---

Declare every namespace prefix a rebuilt part uses. `word/document.xml`, headers, footers, note parts, `comments.xml`, `commentsExtended.xml` and the numbering, styles, settings, font-table and theme parts now derive their `xmlns:*` from the markup they emit instead of a hand-maintained list: each prefix is resolved through one namespace table, falling back to the bindings the source part's root declared, and a prefix that resolves to nothing fails the save with an `UnboundNamespacePrefixError` rather than being written unbound. `mc:Ignorable` comes from the same table, so it names only prefixes the part declares and every declared extension prefix it should. A document carrying content the parser preserves verbatim — a text box under `wne:txbxContent`, a `w16du:dateUtc` revision stamp — no longer saves to a part a consumer refuses to open.
