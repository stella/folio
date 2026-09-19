---
"@stll/folio-core": patch
---

Write page-border art relationship ids in the relationships namespace. `r:id`
on every `w:pgBorders` side, plus `r:topLeft` / `r:topRight` on the top and
`r:bottomLeft` / `r:bottomRight` on the bottom, were written as `w:id`,
`w:topLeft` and so on. Those are different attributes: Word discarded them and
the border art with them, and folio's own prefix-tolerant reader hid it by
reading its own output back. An attribute-less `<w:docGrid/>` is written back
too, rather than dropped for having nothing to say. Nine pairs leave the
container survival baseline.
