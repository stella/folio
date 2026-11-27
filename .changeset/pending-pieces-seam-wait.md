---
"@stll/docx-core": patch
---

Rejecting an insertion made inside another pending insertion now treats that insertion's pieces as the pieces its seam joins, so new content put beside the seam waits on it as it does once the pieces are accepted. Accepting that content keeps the pieces apart, and range markers kept at the seam stay beside it, whichever revisions are resolved first.
