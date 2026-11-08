---
"@stll/folio-core": patch
"@stll/docx-core": patch
---

A canonical session normalizes list rendering after an edit only from the last checkpoint before the first paragraph the edit changed, and only until the list counters match the earlier run again, instead of over every paragraph of the story. `paragraphPropertySnapshotToken` names the snapshot registration a paragraph carries, so a pass can tell that registering again would change nothing.
