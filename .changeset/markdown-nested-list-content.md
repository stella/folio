---
"@stll/folio-core": patch
---

Fixed two markdown round-trip bugs: `toMarkdown` indented a nested list item by a fixed two spaces regardless of its parent's marker width, so a child under an ordered parent (`1. `/`10. ` needs 3/4 spaces, not 2) reimported one level flatter; and `fromMarkdown` silently dropped a table (or a code block/blockquote) nested inside a list item. Such content now survives as a following block, and `fromMarkdown` reports the flattening on `document.warnings` instead of dropping it silently.
