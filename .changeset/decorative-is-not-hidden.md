---
"@stll/docx-core": minor
"@stll/folio-core": patch
---

Read and write a decorative image as the extension Word writes, and keep `hidden` a separate fact. `Image.decorative` was read from a `@decorative` attribute `CT_NonVisualDrawingProps` does not have (no file in the public corpus writes one), and written back as `hidden="1"`, which says the drawing is not displayed — so a decorative image became a hidden one, and re-parsed as neither. It now round-trips through `wp:docPr`'s `{C183D7F6-B498-43B3-948B-1728B52AA6E4}` extension, `Image.hidden` carries `@hidden` on its own and is written identically for inline and anchored drawings, and `Image.docPrExtensions` keeps the other `a:ext` entries of the same list verbatim and in order rather than dropping them. All three survive the editor round trip.
