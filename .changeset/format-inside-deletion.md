---
"@stll/docx-core": patch
---

A direct format lying inside a concurrent direct deletion now leaves the same runs in either order. A sequenced direct format records the run it cut (`cutRun` in its `setRunProps` effect), and a deletion rebased over it joins that run's pieces again with a seam join at its point when the run reaches past both ends of the deletion. The operation schema is now version 11.
