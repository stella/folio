---
"@stll/folio-core": patch
---

Read bounded DOCX archive entries through JSZip's platform-neutral internal stream instead of Node streams, so `ensureParaIds` and `loadDocxArchive` work in browsers and web workers.
