---
"@stll/docx-core": minor
"@stll/folio-core": minor
"@stll/folio-cli": patch
---

Read XML with `@stll/docx-core/xml`, folio's single-pass reader, everywhere: DOCX parsing, package and schema validation, numeric identity checks, preserved-markup checks and the CLI's package checks. `readXml` states its conformance (`ooxml`, `well-formed` or `tolerant`), `wellFormedXmlIssue` replaces separate validation passes, and resource limits and namespace resolution move to `@stll/docx-core/xml/limits` and `@stll/docx-core/xml/namespaces`. Markup that is not well-formed (mismatched or unclosed tags, attributes without quoted values, a DOCTYPE) is refused rather than repaired, and validators also refuse `<` in attribute values, `]]>` in text, `--` in comments and a misplaced XML declaration. `parseXmlWithFastXmlParser` is removed, and fast-xml-parser is no longer a dependency.
