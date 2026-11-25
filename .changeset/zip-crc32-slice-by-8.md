---
"@stll/docx-core": patch
---

Opening a document in a browser checks each zip entry's CRC-32 eight bytes at a time, so the check no longer costs tens of milliseconds per open.
