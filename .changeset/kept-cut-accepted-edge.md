---
"@stll/docx-core": patch
---

Resolving accepted and rejected revisions in either order places kept range markers alike: a pending insertion at the edge of an accepted insertion's content keeps its cut inside that content, and text inserted inside a rejected link takes the link's kept markers into its piece when accepted later.
