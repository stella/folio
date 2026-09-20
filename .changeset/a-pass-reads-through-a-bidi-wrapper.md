---
"@stll/folio-core": patch
---

Read through a bidirectional wrapper in the save-side resource census and the
rendered-page-break detector. `w:bdo` and `w:dir` are transparent, and both
passes stopped at one: a hyperlink authored inside a wrapper got no `r:id`,
which is the whole of how an `href` is saved, so the package held a link
pointing nowhere; and a `w:lastRenderedPageBreak` under one was invisible to
the detector that decides where the break is re-emitted.
