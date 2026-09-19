---
"@stll/folio-core": patch
---

Stop `word/styles.xml` growing by a copy on every save. Which styles the original part already defines was decided by scanning its text for `w:styleId="…"`, and that text is the XML spelling of an id while the model holds its decoded value: a single-quoted attribute, an id carrying an escaped character such as `Header &amp; Footer`, or an earlier attribute whose value contains `>` all read as a style the part lacked, so it was appended again each time the document was saved. The part is now read as XML, and a style id is written at most once.
