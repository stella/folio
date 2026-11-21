---
"@stll/docx-core": minor
"@stll/folio-core": patch
---

Add the `containerStart` insertion point under operation schema 17: paragraphs or a table go before the first block of a story, cell, note, text box or block content control whatever that block is, such as an empty content control, direct or tracked, with exact inverses; a tracked paragraph inserted beside a table may now stand before a content control. Older operation envelopes remain explicitly refused. On the canonical session, a paragraph or table inserted before an empty content control that starts its container goes before it, one inserted beside an empty control a table stands next to goes between them, and table operations address a table whose first cell starts with an empty content control.
