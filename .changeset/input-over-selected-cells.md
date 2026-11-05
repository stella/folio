---
"@stll/folio-core": patch
---

In the canonical session, cancelling a composition over selected table cells no longer deletes their text. The composition was refused and the document restored, and then the browser's removal of the proposal landed on the restored cell selection.
