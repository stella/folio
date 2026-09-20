---
"@stll/folio-core": patch
---

Charge a WordprocessingGroup preview to the package preview budget. The group producer stamped its own data-URL prefix, mime type and filename, so the budget never recognized one and a package retained as many group previews as it happened to contain. The producer now builds its image from the table, and a package past the allowance has the preview dropped, keeping the drawing, its page space and the authored XML the package saves from.
