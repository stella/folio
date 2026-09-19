---
"@stll/folio-core": patch
---

Normalise an unbalanced range marker wherever it sits. The model validator walks a paragraph's whole inline tree; the two parse-boundary normalisers walked `paragraph.content` alone, so a `w:commentRangeStart` or a move-range marker inside `w:ins`, `w:hyperlink`, `w:sdt`, `w:bdo` or `w:dir` was judged but never normalised and made `parseDocx` throw on a document Word opens. Both now read the tree through one exhaustive traversal. The per-kind policy is unchanged: a comment range's unmatched half becomes a point `w:commentReference`, a move range's is dropped.
