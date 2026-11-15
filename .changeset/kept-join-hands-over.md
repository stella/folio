---
"@stll/docx-core": patch
---

Accepting the pieces of an inserted link before an insertion made inside the link, and then that insertion, leaves one link, as accepting all of them at once does. Rejecting a bookmark inserted inside the link while the link's pieces are still pending leaves the same result as rejecting it after accepting them: the pieces join around the bookmark. A resolved piece's join, or a rejection's seam, that meets a pending insertion made inside its record passes to that insertion's facing side.
