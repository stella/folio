---
"@stll/folio-core": patch
---

A full repack no longer refuses a document whose `w:headerReference` or `w:footerReference` states a `w:type` outside `ST_HdrFtr`, such as the `odd` some producers write for the default header. The reference-loss guard reads both sides of its comparison through the same parser, so the value the parse boundary normalised is no longer mistaken for a dropped reference. What the guard refuses is unchanged.
