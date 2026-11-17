---
"@stll/folio-cli": minor
---

`folio sanitize` (MCP `sanitize_document`) removes comments with their companion parts, anchors, and comment-only `people.xml` entries, every `w:rsid*` attribute and the settings `w:rsids` table, and on request the identifying document properties; without a destination it reports what it would remove.
