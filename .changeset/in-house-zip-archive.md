---
"@stll/docx-core": minor
"@stll/utf8-boundaries": minor
"@stll/folio-core": minor
"@stll/folio-cli": patch
---

Read and write DOCX packages with `@stll/docx-core/zip`, a bounded ZIP archive built on the platform's DEFLATE, and add `@stll/docx-core/zip/stream` for streamed archives. Reads check every declared size, ratio and name before inflating and stop each entry at its declared size; writes pass unchanged entries through as their original compressed bytes. `writeZipPart`, `writeZipEntry` and the `no-string-zip-entry` lint rule are removed, `RawDocxContent.originalZip` is a `ZipArchive`, and JSZip is no longer a dependency.
