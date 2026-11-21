---
"@stll/docx-core": patch
---

A paragraph whose mark ends a section can be deleted right before a table, directly or tracked. Its section break goes with it, removed explicitly with an exact inverse, so its section and the next one become one under the later section's properties; accepting the tracked deletion gives the direct result and rejecting it restores the paragraph and its section.
