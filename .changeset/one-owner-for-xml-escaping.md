---
"@stll/docx-core": minor
"@stll/folio-core": patch
---

Give XML escaping one owner, and make its output always well-formed. `@stll/docx-core` now exports `escapeXmlText` and `escapeXmlAttribute`: six hand-rolled escapers disagreed about the characters that matter, so a value could leave folio as markup Word refuses to open, or come back changed. Both functions drop the characters XML 1.0 §2.2 forbids (the C0 controls outside tab/LF/CR, U+FFFE, U+FFFF, unpaired surrogates), which cannot be escaped into a document either. The attribute form writes tab, LF and CR as character references, because §3.3.3 has every conformant reader flatten a literal one to a space; the text form does the same for CR, which §2.11 would otherwise rewrite to LF. `sanitizeXmlCharacters` applies the same rule at an input boundary, where the value can still be reported.
