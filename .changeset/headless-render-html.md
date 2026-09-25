---
"@stll/folio-core": minor
---

Share the headless render path: `buildDocxDisplayList` lays a package out and builds its display list for any backend, `writeDisplayListPdf` writes one, `exportDocxToPdf` accepts `pages` to write a subset, `selectDisplayPages` narrows a display list and re-indexes its links and outline, `renderDisplayListToHtml` serializes the DOM backend's pages to a standalone HTML document with inlined binaries (the DOM backend gains `binaryUrls: "dataUrl"`), and `createFontsourceFaces` routes `@fontsource` faces to the measurer, the PDF writer, and matching `@font-face` rules from caller-supplied file access.
