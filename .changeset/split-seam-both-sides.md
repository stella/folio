---
"@stll/docx-core": patch
---

Rejecting a paragraph mark whose split left no piece of the record it cut on one side no longer keeps that seam on an insertion standing there. So a boundary kept from that insertion stays beside the next piece whether the run's revisions are rejected at once or one at a time.
