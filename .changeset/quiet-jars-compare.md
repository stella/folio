---
"@stll/folio-core": minor
---

Add `compareDocx(base, target, { author, timestamp })`: a deterministic
two-document compare returning the base package with tracked changes that
accept back to the target and reject back to the base, plus a JSON change list
(`insert`, `delete`, `replace`, `move`, `format`, `table-row-insert`,
`table-row-delete`). Header, footer, footnote, and endnote stories are reported
as unsupported rather than silently skipped.

The call reads no clock and no random source, so the same inputs give
byte-identical output. Supporting that, `FolioRevisionStamp` lets any apply
batch pin its revision date and id seed, and `FolioAIBlock.table` records the
block's enclosing table cell.
