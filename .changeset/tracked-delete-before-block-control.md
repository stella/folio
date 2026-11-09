---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

A tracked or suggested deletion of a paragraph that a block content control, custom XML or a preserved block follows now applies in the canonical session instead of being refused. The paragraph's words and its own mark are recorded as deleted; accepting removes the paragraph and rejecting gives it back. Accepting a deleted mark there removes a paragraph left with nothing in it, keeps one that still holds content, and runs content that accepting will take away into the content control's first paragraph, so resolving the words and the mark together or one at a time, in either order, leaves the same document.
