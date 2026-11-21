---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

Opening a package repairs a comment range boundary with no partner between a table's rows or a row's cells, where a reference anchors the comment, and a comment's repeated references (new parse warning `duplicate-comment-reference`); both sessions open such packages.

Comment range boundaries pair only where the pairing is unambiguous, so a lone start never widens a full range; a comment keeps its first full range and reports a later one with its text (new parse warning `dropped-comment-range`); a comment whose only anchor sits between table rows or cells becomes a point at the nearest paragraph.
