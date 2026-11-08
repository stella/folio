---
"@stll/folio-core": patch
---

Rejecting an inserted paragraph mark in front of a suggested paragraph no longer loses the words before it. The joined paragraph used to take every attribute of the paragraph after it, including the marker that makes the whole paragraph a suggested insertion, so its earlier words were left out of saves and removed when the suggestion was rejected. Each paragraph attribute now has its own rule for a join: the suggested paragraph becomes a suggested inserted paragraph mark, and the boundary bookmarks and empty hyperlinks of both paragraphs are kept.
