---
"@stll/docx-core": patch
---

A direct format lying inside a concurrent direct deletion now leaves the same records in either order. A sequenced direct format records the records it cut at both of its ends (`cutRecords` in its `setRunProps` effect), and a deletion rebased over it joins the pieces of each that reaches past both of its ends again with one seam join at its point. The operation schema is now version 11.
