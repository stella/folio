---
"@stll/folio-core": patch
"@stll/folio-react": patch
"@stll/folio-vue": patch
---

Preserve distinct note-reference occurrences through editing and save, and refuse partial occurrence changes before committing them.

Breaking collaboration schema change: attribute schema 12 requires structural note-reference occurrence identities. Snapshots with unattributed references must be re-materialized from the saved DOCX before loading; identities are never inferred from label text.
