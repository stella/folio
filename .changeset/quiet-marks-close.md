---
"@stll/folio-core": patch
---

A comparison covers more of what two documents differ by, and says it in the
markup the format defines for it.

A whole paragraph added or removed now carries its paragraph mark as well as its
runs — `w:pPr/w:rPr/w:ins` and `w:pPr/w:rPr/w:del` — so accepting a deletion
removes the paragraph instead of leaving a blank line, and rejecting an
insertion closes the break instead of leaving an empty one.
Resolving a deleted mark keeps the surviving paragraph's own properties.

A list item that stopped being one is reported, and a paragraph inserted beside a
list item is no longer silently made a further item of that list:
`setBlockParagraphProperties` and the block insertions accept `listLevel: null`,
which clears `w:numPr` the way `styleId: null` already clears `w:pStyle`.

Additions past the base document's last block keep the target's order, so a
paragraph and a table added after it no longer come out table first. Headers and
footers pair by kind and document order when the two packages share no
relationship id, so a comparison of two independently authored documents covers
them instead of reporting them as present on one side only.
