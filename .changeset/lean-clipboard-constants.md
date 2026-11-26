---
"@stll/folio-core": patch
---

Opening a document no longer loads DOMPurify. The editor's paste handling reads the clipboard formats and image files from a module that does not import the paste sanitizer; `@stll/folio-core/utils/clipboard` exports the same API as before.
