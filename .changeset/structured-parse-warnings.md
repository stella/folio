---
"@stll/docx-core": minor
"@stll/folio-core": minor
---

`Document.parseWarnings` reports what the parse boundary normalised, as data: a stable `code` from `PARSE_WARNING_CODES`, the part it happened in, the best position that part can name, and the value folio declined to read. `Document.warnings` is unchanged in shape and is now rendered from that list by one formatter, so the prose and the data cannot disagree. Normalisations that were silent now report: a `w:type` outside `ST_HdrFtr`, a repeated footnote or endnote id, a value outside `ST_OnOff` in either shape, a border with no `w:val`, a `w:comment` with no readable `w:id` (previously read as id 0, which manufactured a duplicate), and a hyperlink naming a relationship its part never defined. Retained warnings are capped per code, with the remainder counted.
