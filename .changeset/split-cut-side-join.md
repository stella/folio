---
"@stll/docx-core": patch
---

Rejecting a tracked change together with a paragraph break inserted inside it joins only the change's own seam, as rejecting them one after another in any order does: the pieces either side of the break no longer join the records beyond them at the break.
