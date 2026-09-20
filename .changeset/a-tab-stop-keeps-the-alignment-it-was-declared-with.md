---
"@stll/docx-core": minor
"@stll/folio-core": minor
---

Generate `TabStopAlignment` from `ST_TabJc`. The union omitted `start` and
`end`, and a stop declared with either left the model entirely, because the
reader needs both a position and an alignment to keep one. The numbering
parser's second, hand-rolled reader for the same enumeration read both as
`left`; it now narrows against the one picklist.
