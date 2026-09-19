---
"@stll/folio-core": patch
---

Save `w:keepNext`, `w:keepLines` and `<w:specVanish/>` when a command sets
them. All three were classified `original-only` in the paragraph write-back
map: an imported value survived through `_originalFormatting`, and a value set
on the paragraph node had no save path at all, so a paragraph with no `w:pPr`
of its own lost it silently. They join `widowControl` as `style-resolved-attr`,
so a commanded value is written and a value that only echoes the paragraph's
style still is not.
