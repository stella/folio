---
"@stll/docx-core": patch
---

Rejecting tracked changes places the boundaries they keep the same way whatever order the changes are resolved in: a boundary at a paragraph's edge no longer stops a paragraph join from rejoining the record beyond it, a one-sided cut keeps outside its piece any boundary that stood beside the removed record, and boundaries kept beside a cut control rejoin with it.
