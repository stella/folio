---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

One shared rule, `finalBreakCarrier` in `@stll/docx-core/ops`, decides which paragraph settles the deletion of a container's final paragraph: the nearest paragraph the batch keeps, or, when the batch deletes every paragraph back to the container's start or a table, the run's first paragraph, whose properties the emptied final paragraph takes.
